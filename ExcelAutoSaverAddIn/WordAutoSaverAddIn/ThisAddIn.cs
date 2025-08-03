using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Timers;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace WordAutoSaverAddIn
{
    public partial class ThisAddIn
    {
        private Timer saveTimer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            saveTimer = new Timer(900000); // 15 minutes
            saveTimer.Elapsed += SaveAllDocuments;
            saveTimer.AutoReset = true;
            saveTimer.Start();
        }

        private void SaveAllDocuments(object sender, ElapsedEventArgs e)
        {
            try
            {
                var app = this.Application;
                foreach (Word.Document doc in app.Documents)
                {
                    if (doc.ReadOnly)
                        continue;

                    if (string.IsNullOrWhiteSpace(doc.Path))
                    {
                        string backupFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "WordAutoBackups");
                        Directory.CreateDirectory(backupFolder);
                        string fileName = $"{Path.GetFileNameWithoutExtension(doc.Name)}_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                        string fullPath = Path.Combine(backupFolder, fileName);
                        doc.SaveAs2(fullPath);
                    }
                    else
                    {
                        doc.Save();
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

        #region VSTO generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}