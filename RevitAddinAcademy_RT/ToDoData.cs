using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAddinAcademy_RT
{
    internal class ToDoData
    {
        public string Task { get; set; }
        public string Status { get; set; }
        public string Display { get; set; }

        public ToDoData(string task, string status)
        {
            Task = task;
            Status = status;
            Display = Task + " : " + Status;
        }

        //Create ToDoData from Display string (i.e. from text file)
        public static ToDoData ReadToDo(string curLine)
        {
            if (curLine.Contains(':'))
            {
                List<string> list = curLine.Split(':').ToList();
                string task = list[0].Trim();
                string status = list[1].Trim();
                ToDoData data = new ToDoData(task, status);
                return data;
            }
            else
            {
                return null;
            }
        }

        public ToDoData CopyData()
        {
            return (ToDoData)this.MemberwiseClone();
        }

        public void RefreshDisplay()
        {
            this.Display = this.Task + " : " + this.Status;
        }
    }
}
