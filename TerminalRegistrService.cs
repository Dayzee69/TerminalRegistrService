using System;
using System.ServiceProcess;
using System.Data.SqlClient;
using System.IO;
using System.Data.Common;
using System.Collections.Generic;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Net.Mail;
using System.Diagnostics;
using System.Linq;

namespace TerminalRegistrService
{
    public partial class TerminalRegistrService : ServiceBase
    {
        
        List<Recipient> recipients = new List<Recipient>(); // лист объектов класса Recipient, который представляет контрагентов 
        List<Payment> payments = new List<Payment>(); // лист объектов класса Payment, который унаследован от Recipient и представляет платежные операции 
        static string strPath = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
        INIManager manager;

        // структура для записи информации всех контрагентов, чтобы отфильтровать в дальнейшем только тех, кто проводил операции за выбранный период
        struct CheckRecipient 
        {
            public string name;
            public string descr;
            public bool exluse;
            
            public CheckRecipient(string name, string descr, bool exluse)
            {
                this.name = name;
                this.descr = descr;
                this.exluse = exluse;

            }
        };
        
        public TerminalRegistrService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Logger.InitLogger();
            Logger.Log.Info("Начало работы службы");
            
            System.Timers.Timer T2 = new System.Timers.Timer();
            T2.Interval = 60000;
            T2.AutoReset = true;
            T2.Enabled = true;
            T2.Start();
            T2.Elapsed += new System.Timers.ElapsedEventHandler(T2_Elapsed);
            
            //Conn();
        }

        private void T2_Elapsed(object sender, EventArgs e)
        {

            try
            {
                string date = DateTime.Now.ToString("HH:mm");
                if (date == "01:31")
                {
                    payments.Clear();
                    recipients.Clear();
                    Conn();
                }

            } catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }

        }

        protected override void OnStop()
        {
            Logger.Log.Info("Завершение работы службы");
        }

        //Метод подключения к БД
        public void Conn()
        {
            manager = new INIManager(strPath + @"\settings.ini");
            SqlConnection sqlConnectionDDS = new SqlConnection(@"Data Source=" + manager.GetPrivateString("Database", "host") + ";Initial Catalog=DirectDataServer;" +
                "User ID=" + manager.GetPrivateString("Database", "login") + ";Password=" + manager.GetPrivateString("Database", "password"));
            SqlConnection sqlConnectionEventWatch = new SqlConnection(@"Data Source=" + manager.GetPrivateString("Database", "host") + ";Initial Catalog=EventWatch2;" +
                "User ID=" + manager.GetPrivateString("Database", "login") + ";Password=" + manager.GetPrivateString("Database", "password"));
            SqlConnection sqlConnectionBazaWRK = new SqlConnection(@"Data Source=" + manager.GetPrivateString("Database", "host") + ";Initial Catalog=BazaWRK;" +
                "User ID=" + manager.GetPrivateString("Database", "login") + ";Password=" + manager.GetPrivateString("Database", "password"));

            try
            {

                sqlConnectionDDS.Open();
                sqlConnectionEventWatch.Open();
                sqlConnectionBazaWRK.Open();

                PaymentQuery(sqlConnectionDDS, sqlConnectionEventWatch, sqlConnectionBazaWRK);

            }
            catch (Exception ex)
            {
                Logger.Log.Error(ex.ToString());
            }
            finally
            {
                sqlConnectionDDS.Close();
                sqlConnectionDDS.Dispose();
                sqlConnectionEventWatch.Close();
                sqlConnectionEventWatch.Dispose();
                sqlConnectionBazaWRK.Close();
                sqlConnectionBazaWRK.Dispose();
            }

        }

        //Метод получения информации из БД 
        private void PaymentQuery(SqlConnection connDDS, SqlConnection connEventWatch, SqlConnection connBazaWRK)
        {
            
            string date1 = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 00:00:00");
            string date2 = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd 23:59:59");

            //string date1 = "2020-09-23 00:00:00";
            //string date2 = "2020-09-23 23:59:59";

            List<CheckRecipient> checkRecipients = new List<CheckRecipient>();

            SqlCommand cmdBazaWRK = connBazaWRK.CreateCommand();
            cmdBazaWRK.CommandTimeout = 0;
            SqlCommand cmdWatch = connEventWatch.CreateCommand();
            cmdWatch.CommandTimeout = 0;
            SqlCommand cmd = connDDS.CreateCommand();
            cmd.CommandTimeout = 0;


            //Выгружается информация о контрагентах из БД DDS, у которых были платежные операции за выбранный период и записываются в список recipients
            cmd.CommandText = "SELECT DISTINCT  t0.RecipientID, t1.InternalName, t1.Description FROM Recipient as t1,Payment as t0 WHERE t0.RecipientID = t1.RecipientID " +
            "and ( PaymentDate >= '" + date1 + "' and PaymentDate <= '" + date2 + "' ) ";
            using (DbDataReader reader = cmd.ExecuteReader())
            {
                if (reader.HasRows)
                {

                    while (reader.Read())
                    {
                        Recipient rec = new Recipient(reader.GetValue(0).ToString(), reader.GetValue(1).ToString());
                        rec.description = reader.GetValue(2).ToString();
                        recipients.Add(rec);
                    }

                }
            }

            for (int i = 0; i < recipients.Count; i++)
            {
                List<string> check = new List<string>();

                //Выгружается информация какие поля(пользовательские кнопки) имеются в таблице PaymentInformation контрагента
                cmd.CommandText = "SELECT COLUMN_NAME FROM information_schema.COLUMNS WHERE TABLE_NAME = '" + recipients[i].internalName + "PaymentInformation'";
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        check.Add(reader.GetValue(0).ToString());
                    }
                }

                //Выгружается информация какие поля(пользовательские кнопки) контрагента имеются в таблице RecipientField
                //Далее записываются только те поля, которые присутствуют в обоих таблицах
                cmd.CommandText = "SELECT InternalName, Name FROM RecipientField WHERE recipientID = " + recipients[i].recipientID;
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        bool kostyl = true;
                        string strHeaders = "№,Дата платежа";
                        while (reader.Read())
                        {
                            if (!reader.GetValue(0).ToString().Contains("summa"))
                            {
                                foreach (var item in check)
                                {
                                    if (item == reader.GetValue(0).ToString())
                                    {
                                        if (kostyl)
                                        {
                                            recipients[i].queryColumns += reader.GetValue(0).ToString();

                                            kostyl = false;
                                        }
                                        else
                                        {
                                            recipients[i].queryColumns += ", " + reader.GetValue(0).ToString();

                                        }
                                        strHeaders += "," + reader.GetValue(1).ToString();
                                    }
                                }

                            }

                        }
                        strHeaders += ",Сумма внесенная,Сумма комиссии,Сумма к зачислению,Вознаграждение";
                        recipients[i].headersColumns = strHeaders.Split(',');
                    }
                }

                //Выгружается информация электронных адресов контрагентов
                cmdBazaWRK.CommandText = "SELECT Email FROM Rcpts WHERE recipientID = " + recipients[i].recipientID;
                using (DbDataReader readerWRK = cmdBazaWRK.ExecuteReader())
                {
                    if (readerWRK.HasRows)
                    {
                        bool kostyl = true;
                        while (readerWRK.Read())
                        {
                            if (kostyl)
                            {
                                recipients[i].email += readerWRK.GetValue(0).ToString();
                                kostyl = false;
                            }
                            else
                            {
                                recipients[i].email += "," + readerWRK.GetValue(0).ToString();
                            }

                        }

                    }
                }


            }

            for (int i = 0; i < recipients.Count; i++)
            {
                cmd.CommandText = "SELECT SessionID, PaymentInformationID FROM Payment WHERE RecipientID = " + recipients[i].recipientID +
                    " and (PaymentDate >= '" + date1 + "' and PaymentDate <= '" + date2 + "')";
                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            Payment payment = new Payment(recipients[i].recipientID, recipients[i].internalName, recipients[i].queryColumns);
                            payment.headersColumns = recipients[i].headersColumns;
                            payment.description = recipients[i].description;
                            payment.session = reader.GetValue(0).ToString();
                            payment.paymentInformationID = reader.GetValue(1).ToString();
                            payments.Add(payment);
                        }

                    }
                }
            }

            for (int i = 0; i < payments.Count; i++)
            {
                cmd.CommandText = "SELECT " + payments[i].queryColumns + " FROM " + payments[i].internalName
                 + "PaymentInformation WHERE PaymentInformationID = " + payments[i].paymentInformationID;

                using (DbDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.HasRows)
                    {

                        while (reader.Read())
                        {
                            string str = payments[i].queryColumns.Replace(" ", "");
                            string[] strArr = payments[i].queryColumns.Split(',');
                            for (int j = 0; j < strArr.Length; j++)
                            {
                                payments[i].columns.Add(strArr[j], reader.GetValue(j).ToString());
                            }

                        }

                    }
                }

                cmdWatch.CommandText = "SELECT Occured, Amount, Comission, Fee, RecipientName FROM PaymentAcceptedEvent WHERE Session = '" +
                payments[i].session + "'";
                using (DbDataReader readerWatch = cmdWatch.ExecuteReader())
                {
                    if (readerWatch.HasRows)
                    {

                        while (readerWatch.Read())
                        {
                            payments[i].paymentDate = readerWatch.GetValue(0).ToString();
                            payments[i].amount = Convert.ToDouble(readerWatch.GetValue(1).ToString());
                            payments[i].comission = Convert.ToDouble(readerWatch.GetValue(2).ToString());
                            payments[i].fee = Convert.ToDouble(readerWatch.GetValue(3).ToString());
                            payments[i].name = readerWatch.GetValue(4).ToString();
                        }

                    }
                }
            }

            cmdBazaWRK.CommandText = "SELECT Recipient, Receiver, Exluse FROM Requisites";
            //Выгружаются данные всех имеющихся контрагентов
            using (DbDataReader readerCheckRecipient = cmdBazaWRK.ExecuteReader())
            {
                if (readerCheckRecipient.HasRows)
                {

                    while (readerCheckRecipient.Read())
                    {
                        CheckRecipient cRec = new CheckRecipient(readerCheckRecipient.GetValue(0).ToString(), readerCheckRecipient.GetValue(1).ToString(),
                                Boolean.Parse(readerCheckRecipient.GetValue(2).ToString()));

                        checkRecipients.Add(cRec);

                    }

                }
            }

            for (int i = 0; i < payments.Count; i++)
            {
                for (int j = 0; j < checkRecipients.Count; j++)
                {
                    if (payments[i].name.Contains(checkRecipients[j].name))
                    {
                        payments[i].descriptionName = checkRecipients[j].descr;
                        payments[i].exluse = checkRecipients[j].exluse;
                    }
                }
            }

            CreateExel();

        }

        //Метод создания Exel файла
        private void CreateExel() 
        {
            for (int i = 0; i < recipients.Count; i++)
            {
                /*
                 * Для создания xlsx файла используется внешняя библиотека EPPlus
                 * Создается объект класса ExelWorksheet, у которого есть метод LoadFromArrays
                 * и принимает список массивов типа object, на основе которых, формирует файл.
                 * Каждый массив object представляет собой строку, элемент массива - столбец данной строки.
                 * Чтобы записать информацию, создается список массивов типа object - CellData
                 */
                if (recipients[i].description.Contains("Гармония"))
                {

                    CreateHarmony(recipients[i]);
                    continue;
                }
                string name = recipients[i].internalName;
                var cellData = new List<object[]>();
                int num = 0;
                int colC = 0;
                string date = payments[0].paymentDate;
                string[] dateArr = date.Split(' ');
                date = dateArr[0];

                if (recipients[i].email != "")
                {

                    for (int j = 0; j < payments.Count; j++)
                    {

                        if (recipients[i].recipientID == payments[j].recipientID)
                        {
                            recipients[i].exluse = payments[j].exluse;
                            ++num;

                            if (num == 1)
                            {

                                object[] objHeader = { "РНКО РИБ:  принятые платежи в пользу " + payments[j].descriptionName };
                                cellData.Add(objHeader);
                                object[] objDate = { "За " + date };
                                cellData.Add(objDate);
                                object[] objDescription = { payments[j].description };
                                cellData.Add(objDescription);
                                object[] objColumnHeaders = payments[j].headersColumns;
                                cellData.Add(objColumnHeaders);
                            }

                            object[] obj = new object[payments[j].columns.Count + 6];
                            colC = payments[j].columns.Count;
                            obj[0] = num;
                            obj[1] = payments[j].paymentDate;
                            int index = 2;
                            foreach (var pair in payments[j].columns)
                            {
                                obj[index] = pair.Value;
                                index++;
                            }
                            obj[index] = payments[j].amount;
                            index++;
                            obj[index] = payments[j].comission;
                            index++;
                            obj[index] = payments[j].amount - payments[j].comission;
                            index++;
                            obj[index] = payments[j].fee;

                            cellData.Add(obj);
                        }

                    }

                    using (ExcelPackage excel = new ExcelPackage())
                    {
                        excel.Workbook.Worksheets.Add(name + " " + date);
                        var worksheet = excel.Workbook.Worksheets[name + " " + date];

                        object[] objItog = new object[] { "Доп.Коммиссия:" };
                        cellData.Add(objItog);

                        object[] objDopComm = new object[] { "Итого:" };
                        cellData.Add(objDopComm);

                        //Оформление шапки, загрузки данных
                        worksheet.Cells["A1:J1"].Merge = true;
                        worksheet.Cells["A2:J2"].Merge = true;
                        worksheet.Cells["A3:J3"].Merge = true;
                        worksheet.Cells[1, 1].LoadFromArrays(cellData);
                        worksheet.Cells.AutoFitColumns();
                        worksheet.Cells[1, 1].Style.Font.Bold = true;
                        worksheet.Cells[3, 1].Style.Font.Bold = true;
                        worksheet.Cells["A1:A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[1, 1].Style.Font.Size = 14;
                        worksheet.Cells[2, 1].Style.Font.Size = 12;
                        worksheet.Cells[3, 1].Style.Font.Size = 12;

                        //Оформление строк Итого, Доп.Комиссия
                        worksheet.Cells["A" + (num + 5) + ":" + Char.ConvertFromUtf32((colC + 3) + 64) + (num + 5)].Merge = true;
                        worksheet.Cells["A" + (num + 6) + ":" + Char.ConvertFromUtf32((colC + 2) + 64) + (num + 6)].Merge = true;
                        worksheet.Cells["A" + (num + 5) + ":" + Char.ConvertFromUtf32((colC + 3) + 64) + (num + 5)].Style.Font.Bold = true;
                        worksheet.Cells["A" + (num + 6) + ":" + Char.ConvertFromUtf32((colC + 2) + 64) + (num + 6)].Style.Font.Bold = true;

                        //Рамка
                        worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Left.Style = ExcelBorderStyle.Thin;

                        //Формулы для строки Итого
                        worksheet.Cells[Char.ConvertFromUtf32((colC + 3) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 3) + 64) + 5 + ":"
                            + Char.ConvertFromUtf32((colC + 3) + 64) + (num + 4) + ")";
                        worksheet.Cells[Char.ConvertFromUtf32((colC + 4) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 4) + 64) + 5 + ":"
                            + Char.ConvertFromUtf32((colC + 4) + 64) + (num + 4) + ")";
                        worksheet.Cells[Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 6) + 64) + 5 + ":"
                            + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 4) + ")";

                        //Формула для строки Доп.Комиссии
                        if (!recipients[i].exluse)
                        {
                            worksheet.Cells[Char.ConvertFromUtf32((colC + 4) + 64) + (num + 5)].Formula = "=IF(" + Char.ConvertFromUtf32((colC + 5) + 64) + (num + 5) + "<10000,20,0)";
                            worksheet.Cells[Char.ConvertFromUtf32((colC + 4) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 4) + 64) + 5 + ":"
                                + Char.ConvertFromUtf32((colC + 4) + 64) + (num + 4) + ")+20";

                        }

                        //Формула для строки Итого 
                        worksheet.Cells[Char.ConvertFromUtf32((colC + 5) + 64) + (num + 6)].Formula = Char.ConvertFromUtf32((colC + 3) + 64) + (num + 6) + "-"
                            + Char.ConvertFromUtf32((colC + 4) + 64) + (num + 6);

                        DirectoryInfo dirInfo = new DirectoryInfo(strPath);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();  
                        }
                        dirInfo.CreateSubdirectory("Registers");
                        FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + " " + date + ".xlsx");
                        excel.SaveAs(excelFile);

                    }
                    Logger.Log.Info("Создан реестр " + name);
                    SendMail(recipients[i].email, name + " " + date);

                }

            }

        }

        private void CreateHarmony(Recipient recipient)
        {

            string name = recipient.internalName;
            var cellData = new List<object[]>();
            int num = 0;
            int colC = 0;
            string date = payments[0].paymentDate;
            string[] dateArr = date.Split(' ');
            date = dateArr[0];
            string checkName = "";
            List<Payment> harmony = new List<Payment>();

            for (int j = 0; j < payments.Count; j++)
            {
                if (payments[j].name.Contains("Гармония"))
                {
                    harmony.Add(payments[j]);
                }
            }

            var sortedpayments = from u in harmony
                              orderby u.name
                              select u;

            if (sortedpayments.Count() !=0)
            {
                checkName = sortedpayments.ElementAt(0).name;
                object[] objHeader = { "РНКО РИБ:  принятые платежи в пользу " + sortedpayments.ElementAt(0).descriptionName };
                cellData.Add(objHeader);
                object[] objDate = { "За " + date };
                cellData.Add(objDate);
                object[] objDescription = { sortedpayments.ElementAt(0).name };
                cellData.Add(objDescription);
                object[] objColumnHeaders = sortedpayments.ElementAt(0).headersColumns;
                cellData.Add(objColumnHeaders);
            }

            for (int i = 0; i < sortedpayments.Count(); i++)
            {
                recipient.exluse = sortedpayments.ElementAt(i).exluse;
                ++num;

                if (checkName != sortedpayments.ElementAt(i).name) 
                {
                    num--;
                    using (ExcelPackage excel = new ExcelPackage())
                    {
                        excel.Workbook.Worksheets.Add(name + " " + date);
                        var worksheet = excel.Workbook.Worksheets[name + " " + date];

                        object[] objItog = new object[] { "Доп.Коммиссия:" };
                        cellData.Add(objItog);

                        object[] objDopComm = new object[] { "Итого:" };
                        cellData.Add(objDopComm);

                        //Оформление шапки, загрузки данных
                        worksheet.Cells["A1:J1"].Merge = true;
                        worksheet.Cells["A2:J2"].Merge = true;
                        worksheet.Cells["A3:J3"].Merge = true;
                        worksheet.Cells[1, 1].LoadFromArrays(cellData);
                        worksheet.Cells.AutoFitColumns();
                        worksheet.Cells[1, 1].Style.Font.Bold = true;
                        worksheet.Cells[3, 1].Style.Font.Bold = true;
                        worksheet.Cells["A1:A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[1, 1].Style.Font.Size = 14;
                        worksheet.Cells[2, 1].Style.Font.Size = 12;
                        worksheet.Cells[3, 1].Style.Font.Size = 12;

                        //Оформление строк Итого, Доп.Комиссия
                        worksheet.Cells["A" + (num + 5) + ":" + Char.ConvertFromUtf32((colC + 3) + 64) + (num + 5)].Merge = true;
                        worksheet.Cells["A" + (num + 6) + ":" + Char.ConvertFromUtf32((colC + 2) + 64) + (num + 6)].Merge = true;
                        worksheet.Cells["A" + (num + 5) + ":" + Char.ConvertFromUtf32((colC + 3) + 64) + (num + 5)].Style.Font.Bold = true;
                        worksheet.Cells["A" + (num + 6) + ":" + Char.ConvertFromUtf32((colC + 2) + 64) + (num + 6)].Style.Font.Bold = true;

                        //Рамка
                        worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Left.Style = ExcelBorderStyle.Thin;

                        //Формулы для строки Итого
                        worksheet.Cells[Char.ConvertFromUtf32((colC + 3) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 3) + 64) + 5 + ":"
                            + Char.ConvertFromUtf32((colC + 3) + 64) + (num + 4) + ")";
                        worksheet.Cells[Char.ConvertFromUtf32((colC + 4) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 4) + 64) + 5 + ":"
                            + Char.ConvertFromUtf32((colC + 4) + 64) + (num + 4) + ")";
                        worksheet.Cells[Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 6) + 64) + 5 + ":"
                            + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 4) + ")";

                        //Формула для строки Доп.Комиссии
                        if (!recipient.exluse)
                        {
                            worksheet.Cells[Char.ConvertFromUtf32((colC + 4) + 64) + (num + 5)].Formula = "=IF(" + Char.ConvertFromUtf32((colC + 5) + 64) + (num + 5) + "<10000,20,0)";
                            worksheet.Cells[Char.ConvertFromUtf32((colC + 4) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 4) + 64) + 5 + ":"
                                + Char.ConvertFromUtf32((colC + 4) + 64) + (num + 4) + ")+20";

                        }

                        //Формула для строки Итого 
                        worksheet.Cells[Char.ConvertFromUtf32((colC + 5) + 64) + (num + 6)].Formula = Char.ConvertFromUtf32((colC + 3) + 64) + (num + 6) + "-"
                            + Char.ConvertFromUtf32((colC + 4) + 64) + (num + 6);
                        
                        DirectoryInfo dirInfo = new DirectoryInfo(strPath);
                        if (!dirInfo.Exists)
                        {
                            dirInfo.Create();
                        }
                        dirInfo.CreateSubdirectory("Registers");
                        if (checkName.Contains("Юность"))
                        {
                            FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_unost " + date + ".xlsx");
                            excel.SaveAs(excelFile);
                            Logger.Log.Info("Создан реестр " + name + "_unost ");
                            SendMail(recipient.email, name + "_unost " + date);
                        }
                        else if (checkName.Contains("Песч"))
                        {
                            FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_pesch " + date + ".xlsx");
                            excel.SaveAs(excelFile);
                            Logger.Log.Info("Создан реестр " + name + "_pesch ");
                            SendMail(recipient.email, name + "_pesch " + date);
                        }
                        else if (checkName.Contains("Коптево"))
                        {
                            FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_koptevo " + date + ".xlsx");
                            excel.SaveAs(excelFile);
                            Logger.Log.Info("Создан реестр " + name + "_koptevo ");
                            SendMail(recipient.email, name + "_koptevo " + date);
                        }
                        else if (checkName.Contains("Ленин"))
                        {
                            FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_lenin " + date + ".xlsx");
                            excel.SaveAs(excelFile);
                            Logger.Log.Info("Создан реестр " + name + "_lenin ");
                            SendMail(recipient.email, name + "_lenin " + date);
                        }
                        else if (checkName.Contains("Дмитр"))
                        {
                            FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_dmitr " + date + ".xlsx");
                            excel.SaveAs(excelFile);
                            Logger.Log.Info("Создан реестр " + name + "_dmitr ");
                            SendMail(recipient.email, name + "_dmitr " + date);
                        }
                        if (checkName.Contains("Мероп"))
                        {
                            FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + " " + date + ".xlsx");
                            excel.SaveAs(excelFile);
                            Logger.Log.Info("Создан реестр " + name);
                            SendMail(recipient.email, name + " " + date);
                        }
                        
                    }

                    
                    //SendMail(recipients[i].email, name + " " + date);

                    checkName = sortedpayments.ElementAt(i).name;
                    cellData.Clear(); 
                    num = 1;
                    object[] objHeader = { "РНКО РИБ:  принятые платежи в пользу " + sortedpayments.ElementAt(i).descriptionName };
                    cellData.Add(objHeader);
                    object[] objDate = { "За " + date };
                    cellData.Add(objDate);
                    object[] objDescription = { sortedpayments.ElementAt(i).name };
                    cellData.Add(objDescription);
                    object[] objColumnHeaders = sortedpayments.ElementAt(i).headersColumns;
                    cellData.Add(objColumnHeaders);
                }

                object[] obj = new object[sortedpayments.ElementAt(i).columns.Count + 6];
                colC = sortedpayments.ElementAt(i).columns.Count;
                obj[0] = num;
                obj[1] = sortedpayments.ElementAt(i).paymentDate;
                int index = 2;
                foreach (var pair in sortedpayments.ElementAt(i).columns)
                {
                    obj[index] = pair.Value;
                    index++;
                }
                obj[index] = sortedpayments.ElementAt(i).amount;
                index++;
                obj[index] = sortedpayments.ElementAt(i).comission;
                index++;
                obj[index] = sortedpayments.ElementAt(i).amount - sortedpayments.ElementAt(i).comission;
                index++;
                obj[index] = sortedpayments.ElementAt(i).fee;

                cellData.Add(obj);
            }

            using (ExcelPackage excel = new ExcelPackage())
            {

                excel.Workbook.Worksheets.Add(name + " " + date);
                var worksheet = excel.Workbook.Worksheets[name + " " + date];

                object[] objItog = new object[] { "Доп.Коммиссия:" };
                cellData.Add(objItog);

                object[] objDopComm = new object[] { "Итого:" };
                cellData.Add(objDopComm);

                //Оформление шапки, загрузки данных
                worksheet.Cells["A1:J1"].Merge = true;
                worksheet.Cells["A2:J2"].Merge = true;
                worksheet.Cells["A3:J3"].Merge = true;
                worksheet.Cells[1, 1].LoadFromArrays(cellData);
                worksheet.Cells.AutoFitColumns();
                worksheet.Cells[1, 1].Style.Font.Bold = true;
                worksheet.Cells[3, 1].Style.Font.Bold = true;
                worksheet.Cells["A1:A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[1, 1].Style.Font.Size = 14;
                worksheet.Cells[2, 1].Style.Font.Size = 12;
                worksheet.Cells[3, 1].Style.Font.Size = 12;

                //Оформление строк Итого, Доп.Комиссия
                worksheet.Cells["A" + (num + 5) + ":" + Char.ConvertFromUtf32((colC + 3) + 64) + (num + 5)].Merge = true;
                worksheet.Cells["A" + (num + 6) + ":" + Char.ConvertFromUtf32((colC + 2) + 64) + (num + 6)].Merge = true;
                worksheet.Cells["A" + (num + 5) + ":" + Char.ConvertFromUtf32((colC + 3) + 64) + (num + 5)].Style.Font.Bold = true;
                worksheet.Cells["A" + (num + 6) + ":" + Char.ConvertFromUtf32((colC + 2) + 64) + (num + 6)].Style.Font.Bold = true;

                //Рамка
                worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                worksheet.Cells["A4:" + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Style.Border.Left.Style = ExcelBorderStyle.Thin;

                //Формулы для строки Итого
                worksheet.Cells[Char.ConvertFromUtf32((colC + 3) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 3) + 64) + 5 + ":"
                    + Char.ConvertFromUtf32((colC + 3) + 64) + (num + 4) + ")";
                worksheet.Cells[Char.ConvertFromUtf32((colC + 4) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 4) + 64) + 5 + ":"
                    + Char.ConvertFromUtf32((colC + 4) + 64) + (num + 4) + ")";
                worksheet.Cells[Char.ConvertFromUtf32((colC + 6) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 6) + 64) + 5 + ":"
                    + Char.ConvertFromUtf32((colC + 6) + 64) + (num + 4) + ")";

                //Формула для строки Доп.Комиссии
                if (!recipient.exluse)
                {
                    worksheet.Cells[Char.ConvertFromUtf32((colC + 4) + 64) + (num + 5)].Formula = "=IF(" + Char.ConvertFromUtf32((colC + 5) + 64) + (num + 5) + "<10000,20,0)";
                    worksheet.Cells[Char.ConvertFromUtf32((colC + 4) + 64) + (num + 6)].Formula = "=SUM(" + Char.ConvertFromUtf32((colC + 4) + 64) + 5 + ":"
                        + Char.ConvertFromUtf32((colC + 4) + 64) + (num + 4) + ")+20";
                }

                //Формула для строки Итого 
                worksheet.Cells[Char.ConvertFromUtf32((colC + 5) + 64) + (num + 6)].Formula = Char.ConvertFromUtf32((colC + 3) + 64) + (num + 6) + "-"
                    + Char.ConvertFromUtf32((colC + 4) + 64) + (num + 6);

                DirectoryInfo dirInfo = new DirectoryInfo(strPath);
                if (!dirInfo.Exists)
                {
                    dirInfo.Create();
                }
                dirInfo.CreateSubdirectory("Registers");
                if (checkName.Contains("Юность"))
                {
                    FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_unost " + date + ".xlsx");
                    excel.SaveAs(excelFile);
                    Logger.Log.Info("Создан реестр " + name + "_unost ");
                    SendMail(recipient.email, name + "_unost " + date);
                }
                else if (checkName.Contains("Песч"))
                {
                    FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_pesch " + date + ".xlsx");
                    excel.SaveAs(excelFile);
                    Logger.Log.Info("Создан реестр " + name + "_pesch ");
                    SendMail(recipient.email, name + "_pesch " + date);
                }
                else if (checkName.Contains("Коптево"))
                {
                    FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_koptevo " + date + ".xlsx");
                    excel.SaveAs(excelFile);
                    Logger.Log.Info("Создан реестр " + name + "_koptevo ");
                    SendMail(recipient.email, name + "_koptevo " + date);
                }
                else if (checkName.Contains("Ленин"))
                {
                    FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_lenin " + date + ".xlsx");
                    excel.SaveAs(excelFile);
                    Logger.Log.Info("Создан реестр " + name + "_lenin ");
                    SendMail(recipient.email, name + "_lenin " + date);
                }
                else if (checkName.Contains("Дмитр"))
                {
                    FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + "_dmitr " + date + ".xlsx");
                    excel.SaveAs(excelFile);
                    Logger.Log.Info("Создан реестр " + name + "_dmitr ");
                    SendMail(recipient.email, name + "_dmitr " + date);
                }
                if (checkName.Contains("Мероп"))
                {
                    FileInfo excelFile = new FileInfo(strPath + @"\Registers\Register " + name + " " + date + ".xlsx");
                    excel.SaveAs(excelFile);
                    Logger.Log.Info("Создан реестр " + name);
                    SendMail(recipient.email, name + " " + date);
                }
            }
        }

        //Метод отправки сообщения
        private void SendMail(string email, string fileName) 
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(manager.GetPrivateString("SMTP", "host"));

                //email
                mail.From = new MailAddress(manager.GetPrivateString("SMTP", "from"));
                mail.To.Add(email);
                //mail.To.Add("glotov@ribank.ru");
                mail.Subject = "Реестр платежей по терминалам";
                mail.Body += "Реестр во вложении.";
                mail.Attachments.Add(new Attachment(strPath + @"\Registers\Register " + fileName + ".xlsx"));


                SmtpServer.Port = Int32.Parse(manager.GetPrivateString("SMTP", "port"));
                SmtpServer.Credentials = new System.Net.NetworkCredential(manager.GetPrivateString("SMTP", "login"), manager.GetPrivateString("SMTP", "password"));

                SmtpServer.EnableSsl = false;
                SmtpServer.Send(mail);
                Logger.Log.Info("Отправлено " + fileName);
            }
            catch (Exception ex) 
            {
                Logger.Log.Error(ex.ToString());
            }

        }

    }
}
