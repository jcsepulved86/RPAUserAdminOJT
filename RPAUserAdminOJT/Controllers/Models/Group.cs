using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAUserAdminOJT.Controllers.Models
{
    public class Group
    {
        public string Name { set; get; }
        public string Description { set; get; }
        public string DistinguishedName { set; get; }
        public string SamAccountName { set; get; }
        public string StructuralObjectClass { set; get; }
    }
}
