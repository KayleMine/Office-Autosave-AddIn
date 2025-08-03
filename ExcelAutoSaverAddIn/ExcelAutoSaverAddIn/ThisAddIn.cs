using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Timers;
using System.IO;

namespace ExcelAutoSaverAddIn
{
    public partial class ThisAddIn
    {
        private Timer saveTimer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            saveTimer = new Timer(900000); // 15 мин
            saveTimer.Elapsed += SaveAllWorkbooks;
            saveTimer.AutoReset = true;
            saveTimer.Start();
        }
        private void SaveAllWorkbooks(object sender, ElapsedEventArgs e)
        {
            try
            {
                var app = this.Application;
                foreach (Excel.Workbook wb in app.Workbooks)
                {
                    if (wb.ReadOnly)
                        continue;

                    if (string.IsNullOrWhiteSpace(wb.Path))
                    {
                        string backupFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ExcelAutoBackups");

                        Directory.CreateDirectory(backupFolder);

                        string fileName = $"{Path.GetFileNameWithoutExtension(wb.Name)}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                        string fullPath = Path.Combine(backupFolder, fileName);

                        wb.SaveAs(fullPath);
                    }
                    else
                    {
                        wb.Save();
                    }
                }
            }
            catch (Exception){}
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            saveTimer?.Stop();
            saveTimer?.Dispose();
        }

        #region Код, автоматически созданный VSTO

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
