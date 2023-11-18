using System;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Net;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Globalization;
using Microsoft.Win32;
using System.Diagnostics;

namespace KodinMyynti
{
    public partial class Form1 : Form
    {
        private System.Threading.Timer dailyTimer;
        private string format = "dd.MM.yyyy";
        private string rootPath = AppDomain.CurrentDomain.BaseDirectory;
        private DateTime today = DateTime.Today;
        private string k = "Kyllä";
        private Color colRed = Color.Red;
        private Color colGreen = Color.Green;
        private int stopday;
        public Form1()
        {
            InitializeComponent();

            AddToStartup();
            this.Load += Form1_Load;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DoWork();

            DateTime now = DateTime.Now;
            DateTime scheduledTime = new DateTime(now.Year, now.Month, now.Day, 23, 00, 0);
            if (now > scheduledTime)
            {
                scheduledTime = scheduledTime.AddDays(1);
            }
            double interval = (scheduledTime - now).TotalMilliseconds;

            dailyTimer = new System.Threading.Timer(TimerCallback, null, Convert.ToInt32(interval), Timeout.Infinite);
        }

        private void TimerCallback (object state) 
        {
            DoWork();

            dailyTimer.Change((int)TimeSpan.FromDays(1).TotalMilliseconds, Timeout.Infinite);
        }

        private void AddToStartup()
        {
            try
            {
                string fileNameExe = "KodinMyynti.exe";               
                string appName = "KodinMyynti"; // Замените на имя вашей программы
                string appPath = Path.Combine(rootPath, fileNameExe); // Замените на путь к исполняемому файлу вашей программы
                                                                             // Получаем текущий ключ автозагрузки
                RegistryKey key = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);

                // Добавляем программу в автозагрузку
                key.SetValue(appName, appPath);

                // Закрываем ключ
                key.Close();
            }
            catch (Exception ex)
            {
                // Обработка ошибок, если не удалось добавить программу в автозагрузку
                Debug.WriteLine("Error adding to startup: " + ex.Message);
            }
        }

        private void DoWork()
        {
            this.BeginInvoke((MethodInvoker)delegate 
            {
                ImportExcelToDB(dataGridView1, -1); 
                this.Hide(); 
            });
        }

        private void ImportExcelToDB(DataGridView db, int dd)
        {
            int a = 0;
            stopday = dd;
            DateTime yesterday = DateTime.Today.AddDays(dd);
            string yesterdayDate = yesterday.ToString(format);
            string fileNameOpen = "Asunnot_" + yesterdayDate + ".xlsx";
            string filePathOpen = Path.Combine(rootPath, fileNameOpen);

            if (File.Exists(filePathOpen))
            {


                using (ExcelPackage package = new ExcelPackage(new System.IO.FileInfo(filePathOpen)))
                {
                    // If you use EPPlus in a noncommercial context
                    // according to the Polyform Noncommercial license:
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                    int rowCount = worksheet.Dimension.Rows;
                    int columnCount = worksheet.Dimension.Columns;
                    
                    db.Rows.Clear();
                    db.Columns.Clear();


                    for (int col = 1; col <= columnCount; col++)
                    {
                        string headerText = worksheet.Cells[1, col].Text;
                        var columnWidth = worksheet.Column(col).Width;

                        var column = new DataGridViewTextBoxColumn
                        {
                            HeaderText = headerText,
                            Width = Convert.ToInt32(columnWidth)
                        };
                        db.Columns.Add(column);
                    }


                    for (int row = 1; row <= rowCount; row++)
                    {
                        db.Rows.Add();
                        for (int col = 1; col <= columnCount; col++)
                        {
                            ExcelRange cellOpen = worksheet.Cells[row, col];
                            object cellValue = cellOpen.Text;

                            if (!string.IsNullOrWhiteSpace(cellValue.ToString()))
                            {
                                a = 0;
                                if (DateTime.TryParse(cellValue.ToString(), out DateTime dateValue))
                                {
                                    // Отображаем только день, месяц и год, исключая время
                                    db.Rows[row - 1].Cells[col - 1].Value = dateValue.Date.ToString(format);
                                }
                                else
                                {
                                    db.Rows[row - 1].Cells[col - 1].Value = cellValue;
                                }
                            }
                            else
                            {
                                a++;
                                if (a == 28)
                                {
                                    rowCount = row;
                                    row = rowCount + 1;
                                }
                            }

                            // Переносим форматирование                         
                            string argbColor = cellOpen.Style.Fill.BackgroundColor.Rgb;
                            if (!string.IsNullOrEmpty(argbColor) && argbColor.Length == 8)
                            {
                                Color color = Color.FromArgb(
                                    Convert.ToInt32(argbColor.Substring(0, 2), 16),
                                    Convert.ToInt32(argbColor.Substring(2, 2), 16),
                                    Convert.ToInt32(argbColor.Substring(4, 2), 16),
                                    Convert.ToInt32(argbColor.Substring(6, 2), 16));

                                db.Rows[row - 1].Cells[col - 1].Style.BackColor = color;
                            }

                            argbColor = cellOpen.Style.Font.Color.Rgb;
                            if (!string.IsNullOrEmpty(argbColor) && argbColor.Length == 8)
                            {
                                Color fontColor = Color.FromArgb(
                                    Convert.ToInt32(argbColor.Substring(0, 2), 16),
                                    Convert.ToInt32(argbColor.Substring(2, 2), 16),
                                    Convert.ToInt32(argbColor.Substring(4, 2), 16),
                                    Convert.ToInt32(argbColor.Substring(6, 2), 16));
                                db.Rows[row - 1].Cells[col - 1].Style.ForeColor = fontColor;
                            }

                            // Устанавливаем шрифт (просто имя шрифта)
                            db.Rows[row - 1].Cells[col - 1].Style.Font = new Font(cellOpen.Style.Font.Name, (float)cellOpen.Style.Font.Size);
                        }
                    }
                }
                CheckWebsiteStatus(dataGridView1);
            }
            else
            {
                if (stopday >= -15)
                {
                    ImportExcelToDB(dataGridView1, dd - 1);
                }
            }

            
        }

        private void CheckWebsiteStatus(DataGridView dbase)
        {
            int rowCount = dbase.RowCount;
            string url;

            for (int i = 0; i < rowCount - 1; i++)
            {

                if (dbase.Rows[i].Cells[2].Value == null)
                {
                    if (dbase.Rows[i].Cells[0].Value != null)
                    {
                        string kohdeID = dbase[0, i].Value.ToString();
                        url = "https://www.etuovi.com/kohde/" + kohdeID; // Замените на адрес сайта, который вы хотите проверить

                        try
                        {
                            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                            request.Method = "HEAD"; // Используем HEAD-запрос для получения только заголовков ответа

                            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                            {
                                if (response.StatusCode == HttpStatusCode.Gone)
                                {
                                    //Error 410 = Sold (or changed)
                                    Kyllä(dbase, i);
                                }
                            }
                        }
                        catch (WebException ex)
                        {
                            if (ex.Response is HttpWebResponse webResponse)
                            {
                                if (webResponse.StatusCode == HttpStatusCode.Gone)
                                {
                                    Kyllä(dbase, i);
                                }
                            }
                        }
                    }
                    else
                    {
                        dbase[1, 0].Style.BackColor = colRed;
                    }
                }
                else
                {
                    if (dbase.Rows[i].Cells[2].Value.ToString() != k &&
                         dbase.Rows[i].Cells[2].Value.ToString() != "kyllä" &&
                         dbase.Rows[i].Cells[2].Value.ToString() != "KYLLÄ")
                    {
                        if (dbase.Rows[i].Cells[0].Value != null)
                        {
                            string kohdeID = dbase[0, i].Value.ToString();
                            url = "https://www.etuovi.com/kohde/" + kohdeID; // Замените на адрес сайта, который вы хотите проверить

                            try
                            {
                                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                                request.Method = "HEAD"; // Используем HEAD-запрос для получения только заголовков ответа

                                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                                {
                                    if (response.StatusCode == HttpStatusCode.Gone)
                                    {
                                        //Error 410 = Sold (or changed)
                                        Kyllä(dbase, i);
                                    }
                                }
                            }
                            catch (WebException ex)
                            {
                                if (ex.Response is HttpWebResponse webResponse)
                                {
                                    if (webResponse.StatusCode == HttpStatusCode.Gone)
                                    {
                                        Kyllä(dbase, i);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            ExportDBToExcel(dataGridView1);
        }

        private void Kyllä(DataGridView dgw, int i)
        {
            dgw.Rows[i].Cells[2].Value = k;

            if (dgw.Rows[i].Cells[1].Value != null)
            {

                if (DateTime.TryParseExact(dgw.Rows[i].Cells[1].Value.ToString(), format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime cellDate))
                {
                    if (today <= cellDate)
                    {
                        for (int col = 0; col < dgw.Columns.Count; col++)
                        {
                            dgw[col, i].Style.BackColor = colGreen;
                        }
                    }
                    else
                    {
                        for (int col = 0; col < dgw.Columns.Count; col++)
                        {
                            dgw[col, i].Style.BackColor = colRed;
                        }
                    }
                }
                else
                {
                    dgw[1, i].Style.BackColor = colRed;
                    dgw[1, i].Style.Font = new System.Drawing.Font(dgw.Font, FontStyle.Bold); ;
                }
            }
            else
            {
                dgw[1, i].Style.BackColor = colRed;
            }
        }

        private void ExportDBToExcel(DataGridView db)
        {
            string todayDate = today.ToString(format);
            string fileNameSave = "Asunnot_" + todayDate + ".xlsx";
            string filePathSave = Path.Combine(rootPath, fileNameSave);

            // Создаем новую книгу Excel
            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excel = new ExcelPackage())
            {
                // Добавляем лист Excel
                ExcelWorksheet workSheet = excel.Workbook.Worksheets.Add("Data");

                // Заполняем лист данными из DataGridView
                for (int row = 0; row < db.Rows.Count; row++)
                {
                    for (int col = 0; col < db.Columns.Count; col++)
                    {
                        workSheet.Cells[row + 1, col + 1].Value = db[col, row].Value;

                        // Получите стиль из DataGridView
                        DataGridViewCellStyle cellStyle = db[col, row].InheritedStyle;

                        // Примените форматирование к соответствующей ячейке Excel
                        ExcelRange cellSave = workSheet.Cells[row + 1, col + 1];
                        cellSave.Style.Font.Name = cellStyle.Font.FontFamily.Name;
                        cellSave.Style.Font.Size = cellStyle.Font.Size;
                        cellSave.Style.Font.Bold = cellStyle.Font.Bold;
                        cellSave.Style.Font.Color.SetColor(db.Rows[row].Cells[col].Style.ForeColor);

                        bool isTransparent = cellStyle.BackColor == Color.Transparent || cellStyle.BackColor.A == 0 || cellStyle.BackColor == Color.White;
                        if (isTransparent)
                        {
                            cellSave.Style.Fill.PatternType = ExcelFillStyle.None;
                            cellSave.Style.Fill.BackgroundColor.SetColor(Color.Transparent);
                        }
                        else
                        {
                            cellSave.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cellSave.Style.Fill.BackgroundColor.SetColor(cellStyle.BackColor);
                        }
                                               
                        // И так далее, в зависимости от того, какие параметры форматирования вам нужно скопировать
                    }
                    workSheet.Cells[row + 1, 2].Style.Numberformat.Format = format;
                }

                for (int col = 0; col < db.Columns.Count; col++)
                {
                    // Устанавливаем ширину столбца в Excel равной ширине столбца в DataGridView
                    workSheet.Column(col + 1).Width = Convert.ToDouble(db.Columns[col].Width);
                }

                using (MemoryStream stream = new MemoryStream())
                {
                    excel.SaveAs(stream);
                    File.WriteAllBytes(filePathSave, stream.ToArray());                    
                }
            }
        }
    }
}
