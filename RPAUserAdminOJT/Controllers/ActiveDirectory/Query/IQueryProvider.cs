using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAUserAdminOJT.Controllers.ActiveDirectory.Query
{
    public interface IQueryProvider
    {
        string GetDistinguishedName(string sAMAccountName);
        UserPrincipal Get(string sAMAccountName);
        List<string> Get(string ldapPath, int userId);
        List<string> Get(string ldapPath, string sAMAccountName); 
        List<Models.Group> GetGroups(string sAMAccountName);
        IEnumerable<UserPrincipal> All();
        bool IsMemberOf(string sAMAccountName, string groupName);
        bool Exist(string sAMAccountName);
    }
}
