using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;

namespace TagCloud4
{
    public partial class ThisAddIn
    {
        private CustomTaskPane taskPane;
        private TaskPaneControl control;
        private int OutlookLanguageID;
        //private Categories categories;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            OutlookLanguageID = Application.LanguageSettings.get_LanguageID(Office.MsoAppLanguageID.msoLanguageIDInstall);            
            control = new TaskPaneControl();
            taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(control, "My Categories");
            taskPane.Visible = true;
            control.getTags(Application);
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public CustomTaskPane CustomTaskPane
        {
            get
            {
                return taskPane;
            }
            
        }

        public int Language
        {
            get
            {
                return OutlookLanguageID;
            }
        }

        public TaskPaneControl TaskPaneControl
        {
            get
            {
                return control;
            }
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
