using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportAsExcelFile
{
    internal class Employee
    {
        public Employee(string name, string firstName)
        {
            Name = name;
            FirstName = firstName;
        }

        public Employee(string name, string firstName, DateTime entryTime, DateTime exitTime)
        {
            Name = name;
            FirstName = firstName;
            EntryTime = entryTime;
            ExitTime = exitTime;
        }

        public Employee(int iD, string name, string firstName, DateTime? entryTime, DateTime? exitTime)
        {
            ID = iD;
            Name = name;
            FirstName = firstName;
            EntryTime = entryTime;
            ExitTime = exitTime;
        }

        public int ID { get; set; }=new Random().Next();
        public string Name { get; set; }
        public string FirstName { get; set; }
        public DateTime? EntryTime { get; set; }
        public DateTime? ExitTime { get; set; }

    }
}
