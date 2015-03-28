namespace TagCloud4
{
    partial class FormRegion1 : Microsoft.Office.Tools.Outlook.ImportedFormRegionBase
    {
        private Microsoft.Office.Interop.Outlook.OlkLabel label1;
        private Microsoft.Office.Interop.Outlook._DRecipientControl to;
        private Microsoft.Office.Interop.Outlook.OlkTextBox subject;
        private Microsoft.Office.Interop.Outlook.OlkCommandButton toButton;
        private Microsoft.Office.Interop.Outlook.OlkLabel subjectLabel;
        private Microsoft.Office.Interop.Outlook._DDocSiteControl message;
        private Microsoft.Office.Interop.Outlook._DRecipientControl _RecipientControl1;
        private Microsoft.Office.Interop.Outlook.OlkCommandButton olkCommandButton1;
        private Microsoft.Office.Interop.Outlook.OlkTextBox textBox1;

        public FormRegion1(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            : base(Globals.Factory, formRegion)
        {
            this.FormRegionShowing += new System.EventHandler(this.FormRegion1_FormRegionShowing);
            this.FormRegionClosed += new System.EventHandler(this.FormRegion1_FormRegionClosed);
        }

        protected override void InitializeControls()
        {
            this.label1 = (Microsoft.Office.Interop.Outlook.OlkLabel)GetFormRegionControl("Label1");
            this.to = (Microsoft.Office.Interop.Outlook._DRecipientControl)GetFormRegionControl("To");
            this.subject = (Microsoft.Office.Interop.Outlook.OlkTextBox)GetFormRegionControl("Subject");
            this.toButton = (Microsoft.Office.Interop.Outlook.OlkCommandButton)GetFormRegionControl("ToButton");
            this.subjectLabel = (Microsoft.Office.Interop.Outlook.OlkLabel)GetFormRegionControl("SubjectLabel");
            this.message = (Microsoft.Office.Interop.Outlook._DDocSiteControl)GetFormRegionControl("Message");
            this._RecipientControl1 = (Microsoft.Office.Interop.Outlook._DRecipientControl)GetFormRegionControl("_RecipientControl1");
            this.olkCommandButton1 = (Microsoft.Office.Interop.Outlook.OlkCommandButton)GetFormRegionControl("OlkCommandButton1");
            this.textBox1 = (Microsoft.Office.Interop.Outlook.OlkTextBox)GetFormRegionControl("TextBox1");

        }

        public partial class FormRegion1Factory : Microsoft.Office.Tools.Outlook.IFormRegionFactory
        {
            public event Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler FormRegionInitializing;

            private Microsoft.Office.Tools.Outlook.FormRegionManifest _Manifest;

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public FormRegion1Factory()
            {
                this._Manifest = Globals.Factory.CreateFormRegionManifest();
                this.InitializeManifest();
                this.FormRegionInitializing += new Microsoft.Office.Tools.Outlook.FormRegionInitializingEventHandler(this.FormRegion1Factory_FormRegionInitializing);
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            public Microsoft.Office.Tools.Outlook.FormRegionManifest Manifest
            {
                get
                {
                    return this._Manifest;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.IFormRegion Microsoft.Office.Tools.Outlook.IFormRegionFactory.CreateFormRegion(Microsoft.Office.Interop.Outlook.FormRegion formRegion)
            {
                FormRegion1 form = new FormRegion1(formRegion);
                form.Factory = this;
                return form;
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            byte[] Microsoft.Office.Tools.Outlook.IFormRegionFactory.GetFormRegionStorage(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(FormRegion1));
                return (byte[])resources.GetObject("alex");
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            bool Microsoft.Office.Tools.Outlook.IFormRegionFactory.IsDisplayedForItem(object outlookItem, Microsoft.Office.Interop.Outlook.OlFormRegionMode formRegionMode, Microsoft.Office.Interop.Outlook.OlFormRegionSize formRegionSize)
            {
                if (this.FormRegionInitializing != null)
                {
                    Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs cancelArgs = Globals.Factory.CreateFormRegionInitializingEventArgs(outlookItem, formRegionMode, formRegionSize, false);
                    this.FormRegionInitializing(this, cancelArgs);
                    return !cancelArgs.Cancel;
                }
                else
                {
                    return true;
                }
            }

            [System.Diagnostics.DebuggerNonUserCodeAttribute()]
            Microsoft.Office.Tools.Outlook.FormRegionKindConstants Microsoft.Office.Tools.Outlook.IFormRegionFactory.Kind
            {
                get
                {
                    return Microsoft.Office.Tools.Outlook.FormRegionKindConstants.Ofs;
                }
            }
        }
    }

    partial class WindowFormRegionCollection
    {
        internal FormRegion1 FormRegion1
        {
            get
            {
                foreach (var item in this)
                {
                    if (item.GetType() == typeof(FormRegion1))
                        return (FormRegion1)item;
                }
                return null;
            }
        }
    }
}
