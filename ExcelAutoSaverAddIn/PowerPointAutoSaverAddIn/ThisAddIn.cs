using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Timers;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace PowerPointAutoSaverAddIn
{
    public partial class ThisAddIn
    {
        private Timer saveTimer;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            saveTimer = new Timer(900000); // 15 minutes
            saveTimer.Elapsed += SaveAllPresentations;
            saveTimer.AutoReset = true;
            saveTimer.Start();
        }

        private void SaveAllPresentations(object sender, ElapsedEventArgs e)
        {
            try
            {
                var app = this.Application;
                foreach (PowerPoint.Presentation pres in app.Presentations)
                {
                    if (pres.ReadOnly == Microsoft.Office.Core.MsoTriState.msoTrue)
                        continue;

                    if (string.IsNullOrWhiteSpace(pres.Path))
                    {
                        string backupFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "PowerPointAutoBackups");
                        Directory.CreateDirectory(backupFolder);
                        string fileName = $"{Path.GetFileNameWithoutExtension(pres.Name)}_{DateTime.Now:yyyyMMdd_HHmmss}.pptx";
                        string fullPath = Path.Combine(backupFolder, fileName);
                        pres.SaveAs(fullPath);
                    }
                    else
                    {
                        pres.Save();
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