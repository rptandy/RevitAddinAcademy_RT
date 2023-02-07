using System;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAddinAcademy_RT
{
    public partial class FrmToDo : Form
    {
        //initialize data list
        BindingList<ToDoData> dataList = new BindingList<ToDoData>();

        //set text to use 
        string defaultTxt = "To do";
        string altTxt = "Complete";

        public List<string> ToDoList { get; set; }

        public FrmToDo(string filePath)
        {
            InitializeComponent();

            //Set data source and display for listbox
            listBox1.DataSource = dataList;
            listBox1.DisplayMember = "Display";

            //set text file
            string txtFile = filePath;

            //Read existing items from file
            if (File.Exists(txtFile))
            {
                string[] textFile = File.ReadAllLines(txtFile);
                foreach (string line in textFile)
                {
                    ToDoData data = ToDoData.ReadToDo(line);
                    dataList.Add(data);
                }
            }
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            ChangeStatus();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddItem();
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            MoveUp();
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            MoveDown();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            DeleteItem();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            ToDoList = GetToDoList();
            this.Close();
        }

        public void MoveUp()
        {
            MoveItem(-1);
        }

        public void MoveDown()
        {
            MoveItem(1);
        }

        public void MoveItem(int direction)
        {
            int curIndex = listBox1.SelectedIndex;

            // Checking there is an item selected
            if (listBox1.SelectedItem == null || curIndex <0)
                return;

            //Calculate new index
            int newIndex = curIndex + direction;

            //Check for out of range
            if (newIndex < 0 || newIndex>=listBox1.Items.Count)
                return;

            //Copy data
            ToDoData selected = dataList[curIndex];
            ToDoData copy = selected.CopyData();

            //Remove element
            dataList.RemoveAt(curIndex);

            //Insert at new position
            dataList.Insert(newIndex, copy);

            //Restore selection
            listBox1.SetSelected(newIndex, true);
        }

        public void DeleteItem()
        {
            int curIndex = listBox1.SelectedIndex;

            // Checking there is an item selected
            if (listBox1.SelectedItem == null || curIndex < 0)
                return;

            //Delete data
            dataList.RemoveAt(curIndex);
        }

        public void AddItem()
        {
            string text = textBoxNew.Text.Trim();
            //Checking there is text
            if (text == null || text == "")
                return;

            //Create data
            ToDoData data = new ToDoData(text, defaultTxt);
            //Add to list
            dataList.Add(data);
        }

        public void ChangeStatus()
        {
            ToDoData curData = dataList[listBox1.SelectedIndex];

            //switch status
            if (curData.Status == defaultTxt)
                curData.Status = altTxt;

            else curData.Status = defaultTxt;

            //recalculate display
            curData.RefreshDisplay();

            //refresh displaymember
            listBox1.DisplayMember = null;
            listBox1.DisplayMember = "Display";
        }

        public List<string> GetToDoList()
        {
            List<string> toDo = new List<string>();

            foreach(ToDoData data in dataList)
            {
                toDo.Add(data.Display);
            }

            return toDo;
        }

    }
}
