using System;
using System.Collections.Generic;
using System.Linq;
using System.Resources;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TagCloud4
{
    public partial class FormRegion1
    {
        #region 窗体区域工厂

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass("IPM.Note.Contoso")]
        [Microsoft.Office.Tools.Outlook.FormRegionName("TagCloud4.FormRegion1")]
        public partial class FormRegion1Factory
        {
            private void InitializeManifest()
            {
                ResourceManager resources = new ResourceManager(typeof(FormRegion1));
                this.Manifest.FormRegionType = Microsoft.Office.Tools.Outlook.FormRegionType.Replacement;
                this.Manifest.Title = resources.GetString("Title");
                this.Manifest.FormRegionName = resources.GetString("FormRegionName");
                this.Manifest.Description = resources.GetString("Description");
                this.Manifest.ShowInspectorCompose = true;
                this.Manifest.ShowInspectorRead = true;
                this.Manifest.ShowReadingPane = true;

            }

            // 在初始化窗体区域之前发生。
            // 若要阻止窗体区域出现，请将 e.Cancel 设置为 True。
            // 使用 e.OutlookItem 获取对当前 Outlook 项的引用。
            private void FormRegion1Factory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
            }
        }

        #endregion

        // 在显示窗体区域之前发生。
        // 使用 this.OutlookItem 获取对当前 Outlook 项的引用。
        // 使用 this.OutlookFormRegion 获取对窗体区域的引用。
        private void FormRegion1_FormRegionShowing(object sender, System.EventArgs e)
        {
            //MessageBox.Show("hi");
        }

        // 在关闭窗体区域时发生。
        // 使用 this.OutlookItem 获取对当前 Outlook 项的引用。
        // 使用 this.OutlookFormRegion 获取对窗体区域的引用。
        private void FormRegion1_FormRegionClosed(object sender, System.EventArgs e)
        {
        }
    }
}
