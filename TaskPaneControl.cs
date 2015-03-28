using System;
using System.Timers;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Threading;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace TagCloud4
{
    public partial class TaskPaneControl : UserControl
    {
        List<String> tags = new List<String>();
        List<String> cloudList = new List<String>();
        List<string> cloudListCopy = new List<string>();
        List<string> itemList = new List<string>();
        List<string> checkedList = new List<string>();
        List<string> defaultBigList = new List<string>();
        List<string> TagCopy = new List<string>();
        CheckedListBox checkedListBox2 = new CheckedListBox();
        List<string> clickedList = new List<string>();
        List<CheckBox> checkList = new List<CheckBox>();
        List<string> shownList = new List<string>();
        TextBox text = new TextBox();
        private static System.Timers.Timer atimer = new System.Timers.Timer ();
        private static Hashtable _hashtable = new Hashtable();
        //private int thread_Count = 0;
        Outlook.Application app;
        Outlook.Explorer explorer;
        Outlook.MailItem oldItem = null;
        //draw d;
        //int FolderChange_count = 1;
        int language = Globals.ThisAddIn.Language;
        Boolean notSelect = true;
        Boolean clear_search = false;
        String clicked = null; //the clicked result representation 
        String itemCategory = null;//represent the item's categories
        public TaskPaneControl()
        {
            InitializeComponent();
            
        }

        //delegate void draw();
        
        private void TaskPaneControl_Load(object sender, EventArgs e)
        {
                   
        }



        // Draw CheckBox and TagCloud   
        public void getTags(Outlook.Application application)
        {
            app = application;
            explorer = app.ActiveExplorer();
            explorer.SelectionChange += explorer_SelectionChange;
            Outlook.Categories categories = app.Session.Categories;
            foreach( Outlook.Category category in categories)
                itemList.Add(category.Name);
            
            var tmpitems = new List<string>();
            foreach (string item in itemList) tmpitems.Add(item);
            itemList.Clear();
            itemList.AddRange(tmpitems.OrderBy(i => i).ToArray());

            
            
            //draw CheckList Box and add eventhandler to Check    
            foreach (string item in itemList)
            {
                this.checkedListBox1.Items.Add(item,CheckState.Unchecked);
                checkedListBox1.ForeColor = Color.Black;
                checkedListBox1.Font = new Font("Arial", 11);
            }

            checkedListBox1.ItemCheck += new ItemCheckEventHandler(this.Check_Clicked);
            searchBox.TextChanged +=new EventHandler(this.searchBox_TextChanged);
            //System.Timers.Timer time = new System.Timers.Timer(10);
            //time.Elapsed += new ElapsedEventHandler(time_Elapsed);
            //time.AutoReset = false;
            //time.Enabled = true;         
        }

        //void time_Elapsed(object sender, ElapsedEventArgs e)
        //{
        //    explorer.FolderSwitch += explorer_FolderSwitch;
        //}

        //System.Timers.Timer time2 = new System.Timers.Timer(10);
        //public void explorer_FolderSwitch()
        //{

        //    if (FolderChange_count == 1)
        //    {
        //        time2.Elapsed += new ElapsedEventHandler(Handler_FolderSwitch);
        //        time2.AutoReset = false;
        //        time2.Enabled = true;
        //    }
        //    else
        //    {
        //        time2.Elapsed -= Handler_FolderSwitch;
        //        time2.AutoReset = false;
        //        time2.Enabled = false;
        //        time2.Elapsed += Handler_FolderSwitch;
        //        time2.Enabled = true;
        //    }

        //}
        //private void TableMultiValuedProperties()
        //{
        //    const string categoriesProperty =
        //        "http://schemas.microsoft.com/mapi/string/"
        //        + "{00020329-0000-0000-C000-000000000046}/Keywords";
        //    // Inbox
        //    Outlook.Folder folder =
        //        explorer.CurrentFolder as Outlook.Folder;
        //    // Call GetTable with filter for categories
        //    string filter = "@SQL="
        //        + "Not(" + "\"" + categoriesProperty
        //        + "\"" + " Is Null)";
        //    Outlook.Table table =
        //        folder.GetTable(filter,
        //        Outlook.OlTableContents.olUserItems);
        //    // Add categories column and append type specifier
        //    table.Columns.Add(categoriesProperty + "/0000001F");
        //    while (!table.EndOfTable)
        //    {
        //        Outlook.Row nextRow = table.GetNextRow();
        //        string[] categories =
        //            (string[])nextRow[categoriesProperty + "/0000001F"];
        //        //MessageBox.Show("Subject: " + nextRow["Subject"]);
                
        //        //foreach (string category in categories)
        //        //{
        //        //    update_TagClodeList(category);
        //        //}
        //        Debug.WriteLine("\n");
        //    }
        //}

        //public void Handler_FolderSwitch(object sender, ElapsedEventArgs e)
        //{
        ////    FolderChange_count = 2;
        //    if (!notSelect)
        //    {
        //        if (clear_search)
        //        {
        //            //cloudList = new List<string>(cloudListCopy);
        //            //defaultBigList = new List<string>(TagCopy);
        //            //d = new draw(_drawCloud);
        //            //d.Invoke();
        //            //this.richTextBox1.BeginInvoke(new draw(_drawCloud));
        //            clear_search = false;
        //        }
        //        notSelect = true;
        //     //   FolderChange_count = 1;
        //        return;
        //    }
        //    cloudList.Clear();
        //    tags.Clear();
        //    checkedList.Clear();
        //    _hashtable.Clear();
        //    //throw new NotImplementedException();
        //    for (int i = 0; i < checkedListBox1.Items.Count; i++)
        //    {
        //        //if (this.richTextBox1.InvokeRequired)
        //        //{
        //        //    //this.richTextBox1.BeginInvoke(new draw(_drawCloud));
        //        //}
        //        //else
        //        //{
        //            checkedListBox1.SetItemCheckState(i, CheckState.Unchecked);
        //        //}
        //    }
        //    try
        //    {
        //        TableMultiValuedProperties();
        //        //Outlook.NameSpace nsp = app.GetNamespace("MAPI");
        //        //Outlook.View view = explorer.CurrentFolder.CurrentView;
        //        //Outlook.TableView tableview = view as Outlook.TableView;
        //        //Outlook.Table table = tableview.GetTable();
        //        //if (table.GetRowCount() != 0)
        //        //{

        //        //    while (!table.EndOfTable)
        //        //    {
        //        //        Outlook.Row row = table.GetNextRow();
        //        //        if (row != null)
        //        //        {
        //        //            string entryID = row["EntryID"];
        //        //            object item = nsp.GetItemFromID(entryID);
        //        //            if (item is Outlook.MailItem)
        //        //            {
        //        //                Outlook.MailItem mailitem = item as Outlook.MailItem;
        //        //                if (mailitem.Categories != null)
        //        //                    update_TagClodeList(mailitem.Categories);
        //        //            }
        //        //        }
        //        //    }
        //        //}

        //      //  cloudListCopy = new List<string>(cloudList);
        //     //   Before_UpdateTags();
        //       // TagCopy = new List<string>(defaultBigList);


        //        //if (notSelect)
        //        //{
        //        //    //draw d = new draw(_drawCloud);
        //        //    //d.Invoke();
        //        //    this.richTextBox1.BeginInvoke(new draw(_drawCloud));
        //        //}
        //        //FolderChange_count = 1;
        //    }
        //    catch (Exception ec)
        //    {
        //       // FolderChange_count = 1;
        //    }
        //}

        void explorer_SelectionChange()
        {
          
            if (explorer.Selection.Count >= 1)
            {
                Outlook.Selection selection = explorer.Selection;
               // MessageBox.Show(selection.Count.ToString());
                object item = selection[1];
                if (item is Outlook.MailItem)
                {   
                    Outlook.MailItem mailitem = item as Outlook.MailItem;
                    if (oldItem != null)
                    {
                        if (oldItem.EntryID.Equals(mailitem.EntryID))
                            return;
                    }
                    oldItem = mailitem;
                    itemCategory = mailitem.EntryID;
                }
            }
            return;
        }

        //public delegate void listBox();
    

        //public void Update_checkedboxlist()
        //{
        //    Outlook.Categories categories = app.Session.Categories;
          
        //    itemList.Clear();
        //    foreach (Outlook.Category category in categories)
        //        itemList.Add(category.Name);

        //    if (this.checkedListBox1.InvokeRequired)
        //    {
        //        this.checkedListBox1.BeginInvoke(new listBox(Update_checkedboxlist));
        //    }
        //    else
        //    {
        //        checkedListBox1.Items.Clear();
        //    }
        //    var tmpitems = new List<string>();
        //    foreach (string item in itemList) tmpitems.Add(item);
        //    itemList.Clear();
        //    itemList.AddRange(tmpitems.OrderBy(i => i).ToArray());


        //    //draw CheckList Box and add eventhandler to Check   
        //    try
        //    {
        //        foreach (string item in itemList)
        //        {

        //            //if (this.richTextBox1.InvokeRequired)
        //            //{
        //                //this.checkedListBox1.BeginInvoke(new listBox(Update_checkedboxlist));
        //            //}
        //            //else
        //            //{
        //                this.checkedListBox1.Items.Add(item, CheckState.Unchecked);
        //            //}
        //            checkedListBox1.ForeColor = Color.Black;
        //            checkedListBox1.Font = new Font("Arial", 11);

        //        }
        //    }
        //    catch(Exception ce)
        //    {
        //    }
            
        //}

        //private void TextBox_Click(Object sender, EventArgs e)
        //{
        //    TextBox textBox = (TextBox)sender;
        //    clicked = textBox.Text;
        //    if (tags.Contains(clicked))
        //    {
        //        tags.Remove(clicked);
        //        clicked = null;
        //    }
        //    else
        //        tags.Add(clicked);
        //    HandEvent(1);
        //}


        //varible checkList changed. Check EvnentHandler for CheckBox 
        private void Check_Clicked(Object sender, ItemCheckEventArgs e)
        {
            checkedList.Clear();

            foreach (object obj in checkedListBox1.CheckedItems)
                checkedList.Add(obj.ToString());

            if (e.NewValue == CheckState.Checked)
                checkedList.Add(checkedListBox1.Items[e.Index].ToString());
            else
            {
                if (checkedList.Contains(checkedListBox1.Items[e.Index].ToString()))
                    checkedList.Remove(checkedListBox1.Items[e.Index].ToString());
            }

            HandEvent(2);
        }

        //Handle event: ReDraw TagCloud due to checkList and clicked changed
        private void HandEvent(int type)
        {


            ////////////////First Deal Serach Text///////////////
         //   cloudList.Clear();
        //    _hashtable.Clear();

            notSelect = false;
            String searchTxt = "";
            if (language == 2052)
            {
                if (type == 1)
                {

                    foreach (string item in tags)
                    {
                        if (!checkedList.Contains(item))
                            searchTxt += " 类别:=\"" + item + "\"";
                    }
                }

                foreach (string check in checkedList)
                {
                    searchTxt += " 类别:=\"" + check + "\"";

                }
            }
            if (language == 1033 || language == 2057)
            {
                if (type == 1)
                {

                    foreach (string item in tags)
                    {
                        if (!checkedList.Contains(item))
                            searchTxt += " category:=\"" + item + "\"";
                    }
                }

                foreach (string check in checkedList)
                {
                    searchTxt += " category:=\"" + check + "\"";

                }
            }


            if (searchTxt == "")
            {
                searchTxt = "";
                clear_search = true;
                explorer.ClearSearch();
                //Update_checkedboxlist();

            }
            else
                explorer.Search(searchTxt, Outlook.OlSearchScope.olSearchScopeAllFolders);

            //if (!clear_search)
            //{
            //    if (thread_Count >= 2)
            //    {
            //        atimer.Elapsed -= (OnTimedEvent);
            //        atimer.Stop();
            //    }

            //    atimer.Elapsed += (OnTimedEvent);
            //    atimer.AutoReset = false;
            //    atimer.Enabled = true;
            //}
            //else
            //{
            //    atimer.Elapsed -= (OnTimedEvent);
            //    atimer.Stop();
            //}

            clicked = null;


        }



        //void OnTimedEvent(Object source, ElapsedEventArgs e)
        //{
        //    thread_Count++;
        //    Outlook.TableView tableView;
        //    Outlook.Table table;
        //    Outlook.NameSpace nsp = app.GetNamespace("MAPI");
        //    Stopwatch s = new Stopwatch();
        //    s.Start();
        //    TimeSpan t = TimeSpan.FromSeconds(10);
        //    while (s.Elapsed < t)
        //    {
        //        explorer = app.ActiveExplorer();
        //        tableView = explorer.CurrentView as Outlook.TableView;
        //        if (tableView != null)
        //        {
        //            try
        //            {
        //                table = tableView.GetTable();
        //                if (table.GetRowCount() != 0)
        //                {
        //                    t = TimeSpan.FromSeconds(0);
        //                    while (!table.EndOfTable)
        //                    {
        //                        Outlook.Row row = table.GetNextRow();
        //                        string entryID = row["EntryID"];
        //                        object item = nsp.GetItemFromID(entryID);
        //                        if (item is Outlook.MailItem)
        //                        {
        //                            Outlook.MailItem mailitem = item as Outlook.MailItem;
        //                            if (mailitem.Categories != null)
        //                                update_TagClodeList(mailitem.Categories);
        //                        }
        //                    }

        //                }
        //            }
        //            catch (Exception ce)
        //            {
        //            }
        //        }
        //    }
        //    s.Stop();
        //    thread_Count = 1;
        //    d = new draw(_drawCloud);
        //    d.Invoke();
        //}

        public string getItemID()
        {
            return itemCategory;
        }


        //System.Timers.Timer time3 = new System.Timers.Timer(10);
        //Boolean UpdateTags = false;
        //void Before_UpdateTags()
        //{
        //    if (!UpdateTags)
        //    {
        //        time3.Elapsed += update_Tags;
        //        time3.Enabled = true;
        //        time3.AutoReset = false;
        //    }
        //    else
        //    {
        //        time3.Elapsed -= update_Tags;
        //        time3.Enabled = false;
        //        time3.Elapsed += update_Tags;
        //        time3.Enabled = true;
        //        time3.AutoReset = false;
        //    }
            
        //}


        //void update_Tags(object sender, ElapsedEventArgs e)
        //{
        //    UpdateTags = true;
            
        //    if (_hashtable.Count != 0)
        //    {
        //        defaultBigList = new List<string>(checkedList);
        //        int sum = 0;
        //        int i = 0;
        //        for (i = 0; i < cloudList.Count; i++)
        //        {
        //            if (_hashtable.Contains(cloudList[i]))
        //            {
        //                 var tmp = _hashtable[cloudList[i]];
        //                 if (tmp is int)
        //                 {
        //                     int t = (int)tmp;
        //                     sum += t;
        //                 }
        //            }
        //        }
        //        for (i = 0; i < cloudList.Count; i++)
        //        {
        //            try
        //            {
        //                if (_hashtable.Contains(cloudList[i]))
        //                {
        //                    var tmp2 = _hashtable[cloudList[i]];
        //                    if (tmp2 is int)
        //                    {
        //                        int t2 = (int)tmp2;
        //                        if (t2 > sum / cloudList.Count)
        //                        {
        //                            if (!defaultBigList.Contains(cloudList[i]))
        //                                defaultBigList.Add(cloudList[i]);
        //                        }
        //                    }
        //                }
        //            }
        //            catch (Exception ce)
        //            {
        //            }
        //        }
                
        //    }
        //    UpdateTags = true;
        //}


        //System.Timers.Timer time4 = new System.Timers.Timer(10);
        //Boolean BeforeDraw = false;
        //public void Before_Draw()
        //{
        //    if (!BeforeDraw)
        //    {
        //        time4.Elapsed += _drawCloud;
        //        time4.Enabled = true;
        //        time4.AutoReset = false;
        //    }
        //    else
        //    {
        //        time4.Elapsed -= _drawCloud;
        //        time4.Enabled = false;
        //        time4.Elapsed += _drawCloud;
        //        time4.Enabled = true;
        //        time4.AutoReset = false;
        //    }
        //}

        //Boolean Drawing = false;
        //private void _drawCloud()
        //{
        //    if (Drawing)
        //    {
        //        //How to stop this function
        //        this.cloudList.Clear();
        //        Drawing = true;
        //    }
        //    else
        //        Drawing = true;
            
        //    if (cloudList.Count != 0)
        //    {
        //        List<string> tmp = new List<string>();
        //        foreach (string item in cloudList) tmp.Add(item);
        //        cloudList.Clear();
        //        cloudList.AddRange(tmp.OrderBy(i => i).ToArray());
                
        //    }
        //    if (this.richTextBox1.InvokeRequired)
        //    {
        //        this.richTextBox1.Invoke(d);
        //    }
        //    else
        //    {
        //        this.richTextBox1.Controls.Clear();
        //    }
            
        //    int control_Len = this.richTextBox1.Width;            
        //    int currentHeight = 3;
        //    int currentWidth = 0;
        //    int maxLine = 0;

        //    Before_UpdateTags();
            
        //    for(int i=0; i<cloudList.Count; i++)
        //    {
        //        String content = (String)cloudList[i];
                
        //        text = new TextBox();
        //        text.Name = content;
        //        text.Text = content;
        //        text.AutoSize = true;
        //        text.MouseClick +=new MouseEventHandler(this.TextBox_Click);
        //        //default control
        //        text.Font = new Font("Arial", 8);
        //        text.ReadOnly = true;
        //        text.BackColor = Color.White;
        //        text.ForeColor = Color.BlueViolet;
        //        text.BorderStyle = BorderStyle.None;
               
        //        int len = text.GetPositionFromCharIndex(text.TextLength - 1).X + 10;
        //        //Judge whether it is a top5 category
        //        if (defaultBigList.Contains(content))
        //        {
        //            text.Font = new Font("Arial", 12);
        //            len = (text.Text.Length + 1) * 16;
        //            //len = text.GetPositionFromCharIndex(text.TextLength - 1).X;
        //        }
        //        if (tags.Contains(content))
        //        {
        //            text.ForeColor = Color.White;
        //            text.BackColor = Color.DarkBlue;
        //            text.Font = new Font("Arial", 12);
        //            len = (text.Text.Length + 1) * 16;
        //            //len = text.GetPositionFromCharIndex(text.TextLength - 1).X;
                    
        //        }
        //        currentWidth +=  len ;

        //        if (currentWidth >= control_Len)  //currentwidt great than control_len
        //        {
        //            if (len >= control_Len) //string greater than control_len
        //            {
        //                if ((currentWidth - len) == 0) //if currentPosition is at head, draw in this line
        //                {
        //                    //draw
        //                    text.Size = new System.Drawing.Size(control_Len, text.Size.Height);
        //                    text.Location = new System.Drawing.Point(0, currentHeight);
        //                    maxLine = maxLine > text.Size.Height ? maxLine : text.Size.Height;
        //                    currentHeight = currentHeight + maxLine + 10;
        //                    currentWidth = 0;
        //                }
        //                else //if current position is not at head, draw in next line
        //                {
        //                    text.Size = new System.Drawing.Size(control_Len, text.Size.Height);
        //                    currentHeight = currentHeight + maxLine + 10;
        //                    text.Location = new System.Drawing.Point(0, currentHeight);
        //                    maxLine = text.Size.Height;
        //                    currentHeight = currentHeight + maxLine + 10;
        //                    currentWidth = 0;
        //                }
        //            }
        //            else  //string less than control_len, draw it in next Line
        //            {
        //                text.Size = new System.Drawing.Size(len, text.Size.Height);
        //                text.Location = new System.Drawing.Point(0, currentHeight + maxLine + 5);
        //                currentHeight = currentHeight + maxLine + 10;
        //                currentWidth = len + 20;
        //            }
        //        }
        //        else  //draw in this line
        //        {
        //            text.Size = new System.Drawing.Size(len, text.Size.Height);
        //            text.Location = new System.Drawing.Point(currentWidth-len, currentHeight);
        //            maxLine = maxLine > text.Size.Height ? maxLine : text.Size.Height;
        //            currentWidth += 20;
        //        }
        //        //CheckForIllegalCrossThreadCalls = false;
        //        if (this.richTextBox1.InvokeRequired)
        //        {
        //            this.richTextBox1.Invoke(d);
        //        }
        //        else
        //        {
        //            this.richTextBox1.Controls.Add(text);
        //        }
        //    }
        //    Drawing = false;
        //}



        private void searchBox_TextChanged(object sender, EventArgs e)
        {
           
            String text = this.searchBox.Text;
            shownList.Clear();
            shownList.AddRange(itemList);
            foreach (string check in itemList)
            {
                if (!textContains(check, text))
                {
                    shownList.Remove(check);
                }
                else if(!itemList.Contains(check))
                {
                    shownList.Add(check);
                }
            }
            var items = new List<string>();
            foreach (string item in shownList) items.Add(item);
            shownList.Clear();
            shownList.AddRange(items.OrderBy(i => i).ToArray());
            refreshChecker();
        }

        private void refreshChecker()
        {
            checkedListBox1.Items.Clear();
            foreach (string item in shownList)
                checkedListBox1.Items.Add(item, CheckState.Unchecked);
        }

        private bool textContains(String src, String sub)
        {
            return Regex.Match(src, sub, RegexOptions.IgnoreCase).Success; ;
        }

        private void tagPanel_Paint(object sender, PaintEventArgs e)
        {
        }

        private void TagCloudpanel_Paint(object sender, PaintEventArgs e)
        {
        }

      
        //private void update_TagClodeList(string categories)
        //{
        //    String[] categoryList = categories.Split(new char[] { ',' });

        //    int len = categoryList.Length;
        //    int i = 0;
        //    for (i = 0; i < len; i++)
        //    {
        //        categoryList[i] = categoryList[i].Trim();
        //        string t = categoryList[i];
        //        if (!cloudList.Contains(t))
        //        {
        //            cloudList.Add(t);
        //            _hashtable.Add(t, 0);
        //        }
        //        else
        //        {
        //            var tmp = _hashtable[t];
        //            if (tmp is int)
        //            {
        //                int a = (int)tmp;
        //                a = a + 1;
        //                _hashtable[t] = a;
        //            }

        //        }
                    
        //    }
           
        //    if (clicked != null && !cloudList.Contains(clicked))
        //        cloudList.Add(clicked);
        //}
    
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
       
      
        private void richTextBox1_TextChanged_1(object sender, EventArgs e)
        {
            
        }


        private void splitContainer3_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }         
    }
}
