using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;


namespace TagCloud4
{
    public partial class TagCloud
    {
        //private CustomTaskPane taskPane;

        private void TagCloud_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.CustomTaskPane != null)
            {
                if (Globals.ThisAddIn.CustomTaskPane.Visible == false)
                    Globals.ThisAddIn.CustomTaskPane.Visible = true;
                else
                    Globals.ThisAddIn.CustomTaskPane.Visible = false;
            }
        }
    }
}
