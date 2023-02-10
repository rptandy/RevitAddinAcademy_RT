using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace RevitAddinAcademy_RT
{
    public partial class frmToDo_Sln : Form
    {
        string TodoFilePath = "";
        BindingList<TodoData_Sln> todoDataList = new BindingList<TodoData_Sln>();
        TodoData_Sln currentEdit;

        public frmToDo_Sln(string filePath)
        {
            InitializeComponent();

            lblFilename.Text = Path.GetFileName(filePath);

            string curPath = Path.GetDirectoryName(filePath);
            string curFilename = Path.GetFileNameWithoutExtension(filePath) + "_todo.txt";

            TodoFilePath = curPath + "\\" + curFilename;

            ReadTodoFile();
        }

        private void ReadTodoFile()
        {
            if (File.Exists(TodoFilePath))
            {
                int counter = 0;
                string[] strings = File.ReadAllLines(TodoFilePath);

                foreach(string line in strings)
                {
                    string[] todoData = TodoData_Sln.ParseDisplayString(line);
                    TodoData_Sln curTodo = new TodoData_Sln(counter + 1, todoData[0], todoData[1]);
                    
                    todoDataList.Add(curTodo);
                    counter++;
                }
            }

            ShowData();
        }

        private void ShowData()
        {
            lbxTodo.DataSource = null;
            lbxTodo.DataSource = todoDataList;
            lbxTodo.DisplayMember = "Display";
        }

        private void AddTodoItem(string todoText)
        {
            TodoData_Sln curTodo = new TodoData_Sln(todoDataList.Count + 1, todoText, "To do");
            todoDataList.Add(curTodo);

            WriteTodoFile();
        }

        private void RemoveItem(TodoData_Sln curTodo)
        {
            todoDataList.Remove(curTodo);
            ReorderTodoItems();
            WriteTodoFile();
        }

        private void ReorderTodoItems()
        {
            for(int i = 0; i < todoDataList.Count; i++)
            {
                todoDataList[i].PositionNumber = i + 1;
                todoDataList[i].UpdateDisplayString();
            }

            WriteTodoFile();
        }

        private void WriteTodoFile()
        {
            using (StreamWriter writer = File.CreateText(TodoFilePath))
            {
                foreach(TodoData_Sln curTodo in lbxTodo.Items)
                {
                    curTodo.UpdateDisplayString();
                    writer.WriteLine(curTodo.Display);
                }
            }

            ShowData();
        }

        private void btnAddEdit_Click(object sender, EventArgs e)
        {
            if(currentEdit == null)
            {
                AddTodoItem(tbxAddEdit.Text);
            }
            else
            {
                CompleteEditingItem();
            }

            tbxAddEdit.Text = "";
        }

        private void CompleteEditingItem()
        {
            foreach(TodoData_Sln todo in todoDataList)
            {
                if(todo == currentEdit)
                {
                    todo.Text = tbxAddEdit.Text;
                };
            }

            currentEdit = null;
            lblAddEdit.Text = "Add Item";
            btnAddEdit.Text = "Add Item";

            WriteTodoFile();
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if(lbxTodo.SelectedItems != null)
            {
                TodoData_Sln curTodo = lbxTodo.SelectedItem as TodoData_Sln;
                RemoveItem(curTodo);
            }
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            if (lbxTodo.SelectedItems != null)
            {
                TodoData_Sln curTodo = lbxTodo.SelectedItem as TodoData_Sln;
                StartEditingItem(curTodo);
            }
        }

        private void StartEditingItem(TodoData_Sln curTodo)
        {
            currentEdit = curTodo;
            lblAddEdit.Text = "Update Item";
            btnAddEdit.Text = "Update Item";
            tbxAddEdit.Text = curTodo.Text;

        }

        private void lbxTodo_DoubleClick(object sender, EventArgs e)
        {
            if(lbxTodo.SelectedItems != null)
            {
                TodoData_Sln todo = lbxTodo.SelectedItem as TodoData_Sln;
                FinishItem(todo);
            }
        }

        private void FinishItem(TodoData_Sln todo)
        {
            todo.Display = "Complete";
            WriteTodoFile();
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (lbxTodo.SelectedItems != null)
            {
                TodoData_Sln todo = lbxTodo.SelectedItem as TodoData_Sln;
                MoveItemUp(todo);
            }
        }

        private void MoveItemUp(TodoData_Sln todo)
        {
            for(int i=0; i<todoDataList.Count; i++)
            {
                if(todoDataList[i] == todo)
                {
                    if(i != 0)
                    {
                        todoDataList.RemoveAt(i);
                        todoDataList.Insert(i-1, todo);
                        ReorderTodoItems();
                    }
                }
            }

            WriteTodoFile();
        }

        private void btnDn_Click(object sender, EventArgs e)
        {
            TodoData_Sln todo = lbxTodo.SelectedItem as TodoData_Sln;
            MoveItemDown(todo);
        }

        private void MoveItemDown(TodoData_Sln todo)
        {
            for (int i = 0; i < todoDataList.Count; i++)
            {
                if (todoDataList[i] == todo)
                {
                    if (i < todoDataList.Count - 1)
                    {
                        todoDataList.RemoveAt(i);
                        todoDataList.Insert(i + 1, todo);
                        ReorderTodoItems();
                    }
                }
            }

            WriteTodoFile();
        }
    }
}
