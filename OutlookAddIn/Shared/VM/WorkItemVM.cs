using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAddIn.Shared.VM
{
    public class WorkItemVM
    {
        public string name { get; set; }
        public WorkItemVM(string name)
        {
            this.name = name;
        }
    }
}
