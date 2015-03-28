using System;
using System.Timers;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Drawing;
using System.Text.RegularExpressions;


namespace TagCloud4
{
    partial class Categories
    {
        private List<string> itemCategories = new List<string>();
        private List<string> allCategories = new List<string>();
        Outlook.MailItem item;
        string promptString = "";
        ListBox listbox = new ListBox();
        ToolStripDropDown popup = new ToolStripDropDown();
        System.Timers.Timer time = new System.Timers.Timer(10);

        #region 窗体区域工厂

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("TagCloud4.MyFormRegion")]
        public partial class MyFormRegionFactory
        {
            // 在初始化窗体区域之前发生。
            // 若要阻止窗体区域出现，请将 e.Cancel 设置为 True。
            // 使用 e.OutlookItem 获取对当前 Outlook 项的引用。
            private void MyFormRegionFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {

            }
        }

        #endregion

        // 在显示窗体区域之前发生。
        // 使用 this.OutlookItem 获取对当前 Outlook 项的引用。
        // 使用 this.OutlookFormRegion 获取对窗体区域的引用。
        private void MyFormRegion_FormRegionShowing(object sender, System.EventArgs e)
        {
            this.BorderStyle = System.Windows.Forms.BorderStyle.None;
            popup.Visible = false;
            listbox.Visible = false;
            listbox.Dock = DockStyle.Fill;
            popup.AutoClose = false;
            popup.BackColor = Color.Orange;
            string EntryID = Globals.ThisAddIn.TaskPaneControl.getItemID();
            item = Globals.ThisAddIn.Application.Session.GetItemFromID(EntryID);
            if (item.Categories != null)
                this.textBox1.Text = item.Categories + ",";
            else
                this.textBox1.Text = "";
            this.textBox1.Font = new Font("Arial", 10);
            listbox.BackColor = Color.White;

            listbox.ForeColor = Color.Black;
            
            listbox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.TextChanged += new EventHandler(textBox1_TextChanged);
            this.textBox1.KeyDown += new KeyEventHandler(textBox1_KeyDown);
            listbox.KeyDown += new KeyEventHandler(listbox_KeyDown);
            this.listbox.SelectedIndexChanged += new EventHandler(listbox_SelectedIndexChanged);
            listbox.MouseClick += new MouseEventHandler(Categories_MouseClick);

            this.listbox.TabStop = true;
            this.listbox.TabStopChanged += new EventHandler(listbox_TabStopChanged);
            this.textBox1.LostFocus += new EventHandler(textBox1_LostFocus);


        }

    


        void listbox_TabStopChanged(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
            MessageBox.Show("tab stop changed");
        }

        void textBox1_LostFocus(object sender, EventArgs e)
        {
            ////throw new NotImplementedException();
            if (listbox.Visible == false && !textBox1.Focused && !listbox.Focused)
            {
                SaveItems();
            }
        }

        void Categories_MouseClick(object sender, MouseEventArgs e)
        {
            //throw new NotImplementedException();
            if (listbox.Visible == true && listbox.SelectedIndex != -1)
            {
                string content = this.textBox1.Text;
                //content.Trim();
                String[] categoryList = content.Split(new char[] { ',' });
                int len = categoryList[categoryList.Length - 1].Length;
                this.textBox1.Text = content.Substring(0, content.Length - len);
                this.textBox1.Text += listbox.Items[listbox.SelectedIndex].ToString();
                this.textBox1.Text += ",";
                this.listbox.Visible = false;
                this.popup.Visible = false;
                this.textBox1.Select(textBox1.TextLength, 0);
            }
        }

        void listbox_KeyDown(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                if (listbox.Visible == false) return;
                string content = this.textBox1.Text;
                //content.Trim();
                String[] categoryList = content.Split(new char[] { ',' });
                int len = categoryList[categoryList.Length - 1].Length;
                this.textBox1.Text = content.Substring(0, content.Length - len);
                this.textBox1.Text += listbox.Items[listbox.SelectedIndex].ToString();
                this.textBox1.Text += ",";
                this.listbox.Visible = false;
                this.popup.Visible = false;
                textBox1.Focus();
                this.textBox1.Select(textBox1.TextLength, 0);
            }
            if (e.KeyCode != Keys.Up || e.KeyCode != Keys.Down)
            {
                textBox1.Focus();

            }
        }

        void listbox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private bool textContains(String src, String sub)
        {
            try
            {

                return Regex.Match(src, sub, RegexOptions.IgnoreCase).Success; ;
            }
            catch (Exception ce)
            {
                return false;
            }
        }

        public void SaveItems()
        {
            popup.Visible = false;
            listbox.Visible = false;
            Char isOver;
            String content = this.textBox1.Text;
            if (content != null && content != "")
            {

                content = content.Trim();
                isOver = content[content.Length - 1];
                item.Categories = "";
                int tmpLen = content.Length;
                if (isOver.ToString() == ",")
                    content = content.Substring(0, tmpLen - 1);
                String[] categoryList = content.Split(new char[] { ',' });
                for (int j = 0; j < categoryList.Length; j++)
                    categoryList[j] = categoryList[j].Trim();
                List<string> newCategories = new List<string>();
                newCategories.AddRange(categoryList.Distinct());
                newCategories.Remove("");
                for (int i = 0; i < newCategories.Count; i++)
                {
                    item.Categories += "," + newCategories[i];
                    Outlook.Categories categories =
                                Globals.ThisAddIn.Application.Session.Categories;
                    if (!CategoryExists(newCategories[i]))
                    {
                        try
                        {
                            addToMasterCategoryList(newCategories[i]);
                        }
                        catch (Exception ex)
                        {
                            string mx = ex.Source;
                            MessageBox.Show("Category names cannot contain commas or semicolons.");
                        }
                    }
                }
            }
            else
            {
                item.Categories = null;
            }
            this.textBox1.Text = item.Categories + ",";
            try
            {
                item.Save();
            }
            catch (Exception ce)
            {
            }
            this.textBox1.Select(textBox1.TextLength, 0);
            if (time.Enabled == true)
                time.Elapsed -= time_Elapsed;
            time.Elapsed += time_Elapsed;
            time.AutoReset = false;
            time.Enabled = true;
        }

        void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            //throw new NotImplementedException();

            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
            {
                if (this.listbox.Visible == true)
                {
                    listbox.Focus();
                    listbox_KeyDown(listbox, e);
                    return;
                }
                SaveItems();


            }
            else if (e.KeyCode == Keys.Up)
            {
                //textBox1.Focus();
                listbox.Focus();
                if (listbox.Items.Count == 0) return;
                if (listbox.SelectedIndex != 0)
                    listbox.SelectedIndex--;


            }
            else if (e.KeyCode == Keys.Down)
            {
                listbox.Focus();
                if (listbox.Items.Count == 0) return;

                if (listbox.SelectedIndex <= listbox.Items.Count - 2)
                    listbox.SelectedIndex++;
            }
        }

        void time_Elapsed(object sender, ElapsedEventArgs e)
        {
            // Globals.ThisAddIn.TaskPaneControl.Update_checkedboxlist();
        }

        private void addToMasterCategoryList(string category)
        {

            Outlook.Categories categories = Globals.ThisAddIn.Application.Session.Categories;
            Random random = new Random();
            int tmpRandom = random.Next(26);
            switch (tmpRandom)
            {
                case 0: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorBlack); break;
                case 1: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorBlue); break;
                case 2: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkBlue); break;
                case 3: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkGray); break;
                case 4: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkGreen); break;
                case 5: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkMaroon); break;
                case 6: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkOlive); break;
                case 7: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkOrange); break;
                case 8: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkPeach); break;
                case 9: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkPurple); break;
                case 10: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkRed); break;
                case 11: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkSteel); break;
                case 12: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkTeal); break;
                case 13: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorDarkYellow); break;
                case 14: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorGray); break;
                case 15: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorGreen); break;
                case 16: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorMaroon); break;
                case 17: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorNone); break;
                case 18: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorOlive); break;
                case 19: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorOrange); break;
                case 20: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorPeach); break;
                case 21: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorPurple); break;
                case 22: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorRed); break;
                case 23: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorSteel); break;
                case 24: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorTeal); break;
                case 25: categories.Add(category, Outlook.OlCategoryColor.olCategoryColorYellow); break;
                default:
                    categories.Add(category, Outlook.OlCategoryColor.olCategoryColorYellow);
                    break;
            }


        }

        void textBox1_TextChanged(object sender, EventArgs e)
        {
            //last category which will give promote
            //can change any category but last one give promote
            //if find not in categories then add in to mailitem

            popup.Items.Clear();
            popup.Visible = false;

            Char isOver;
            String content = this.textBox1.Text;
            if (content != "" && content != null)
            {
                isOver = content[content.Length - 1];
                if (isOver.ToString() != ",")
                {
                    content = content.Trim();
                    String[] categoryList = content.Split(new char[] { ',' });
                    promptString = categoryList[categoryList.Length - 1].Trim();
                    Update_Listbox(promptString);
                    if (listbox.Items.Count != 0)
                    {
                        ToolStripControlHost host = new ToolStripControlHost(listbox);
                        popup.Items.Add(host);
                        Point newPoint = new Point();
                        newPoint = this.textBox1.GetPositionFromCharIndex(this.textBox1.TextLength - 1);
                        popup.Show(this.textBox1, newPoint, ToolStripDropDownDirection.AboveRight);
                        popup.Visible = false;
                        popup.Show(this.textBox1, newPoint, ToolStripDropDownDirection.AboveRight);
                        //  listbox.Focus();
                    }
                    else
                    {
                        listbox.Visible = false;
                        popup.Visible = false;
                    }

                }
                else
                {
                    listbox.Visible = false;
                    popup.Visible = false;
                }
            }


        }


        private void Update_Listbox(string prompt)
        {
            listbox.Items.Clear();
            Outlook.Categories categories = Globals.ThisAddIn.Application.Session.Categories;
            int count = 0;
            int lenPrompt = prompt.Length;
            int index = 0;
            List<string> tmpList = new List<string>();
            foreach (Outlook.Category item1 in categories)
            {
                if (textContains(item1.Name, prompt))
                {

                    if (count < 5 && (textContains(item1.Name.Substring(0, lenPrompt), prompt)))
                    {
                        count++;
                        listbox.Items.Add(item1.Name);
                    }
                    else
                    {
                        tmpList.Add(item1.Name);
                    }
                }
            }
            if (count < 5)
            {
                if (tmpList.Count != 0)
                {
                    for (index = 0; index < tmpList.Count; index++)
                    {
                        listbox.Items.Add(tmpList[index]);
                        count++;
                        if (count > 5)
                            break;
                    }
                }
            }


            count = listbox.Items.Count;
            if (count > 0)
                listbox.SelectedIndex = 0;

        }


        private bool CategoryExists(string categoryName)
        {
            try
            {
                //Outlook.Category category =
                //    Globals.ThisAddIn.Application.Session.Categories[categoryName];
                List<string> tmpList = new List<string>();
                foreach (Outlook.Category item in Globals.ThisAddIn.Application.Session.Categories)
                {
                    tmpList.Add(item.Name);
                }

                if (tmpList.Contains(categoryName))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch { return false; }
        }


        // 在关闭窗体区域时发生。
        // 使用 this.OutlookItem 获取对当前 Outlook 项的引用。
        // 使用 this.OutlookFormRegion 获取对窗体区域的引用。
        private void MyFormRegion_FormRegionClosed(object sender, System.EventArgs e)
        {
            listbox.Dispose();
            popup.Dispose();
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }
    }
}
