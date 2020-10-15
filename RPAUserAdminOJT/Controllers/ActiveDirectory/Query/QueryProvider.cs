using System;
using System.Collections.Generic;
using System.Linq;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices;

namespace RPAUserAdminOJT.Controllers.ActiveDirectory.Query
{
    public class QueryProvider
    {

        public static PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, Models.GlobalVar.Domain, Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);

        public QueryProvider(PrincipalContext principalContext)
        {
            principalContext = principalContext ?? throw new ArgumentNullException(nameof(principalContext));
        }

        public IEnumerable<UserPrincipal> All()
        {
            using (var searcher = new PrincipalSearcher(new UserPrincipal(principalContext)))
            {
                IEnumerable<UserPrincipal> users = searcher.FindAll().Select(u => (UserPrincipal)u);
                return users;
            }
        }

        public static UserPrincipal Get(string sAMAccountName)
        {
            return UserPrincipal.FindByIdentity(principalContext, sAMAccountName);
        }

        public List<string> Get(string ldapPath, string sAMAccountName)
        {
            List<string> lstUsers = new List<string>();
            using (var entry = new DirectoryEntry($"LDAP://{ldapPath}"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(sAMAccountName={sAMAccountName}*))";
                    SearchResultCollection results = searcher.FindAll();
                    foreach (SearchResult r in results)
                    {
                        lstUsers.Add(r.Properties["samAccountName"][0].ToString());
                    }
                    return lstUsers;
                }
            }
        }

        public List<string> Get(string ldapPath, int userId)
        {
            List<string> lstUsers = new List<string>();
            using (var entry = new DirectoryEntry($"LDAP://{ldapPath}"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=user)(postOfficeBox={userId}))";
                    SearchResultCollection results = searcher.FindAll();
                    foreach (SearchResult r in results)
                    {
                        lstUsers.Add(r.Properties["samAccountName"][0].ToString());
                    }
                    return lstUsers;
                }
            }
        }

        public string GetDistinguishedName(string sAMAccountName)
        {
            UserPrincipal user = Get(sAMAccountName);
            if (user != null)
            {
                DnParser.DN dn = new DnParser.DN(user.DistinguishedName);
                return dn.Parent.ToString();
            }
            return null;
        }

        public static List<Models.Group> GetGroups(string sAMAccountName)
        {
            List<Models.Group> groups = new List<Models.Group>();
            UserPrincipal user = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, sAMAccountName);
            foreach (GroupPrincipal group in user.GetGroups())
            {
                groups.Add(new Models.Group
                {
                    Name = group.Name,
                    Description = group.Description,
                    DistinguishedName = group.DistinguishedName,
                    SamAccountName = group.SamAccountName,
                    StructuralObjectClass = group.StructuralObjectClass
                });
            }
            return groups;
        }

        public Models.Group GetGroup(string ldapPath, string groupName)
        {
            Models.Group group = new Models.Group();
            using (var entry = new DirectoryEntry($"LDAP://{ldapPath}"))
            {
                using (var searcher = new DirectorySearcher(entry))
                {
                    searcher.Filter = $"(&(objectClass=group)(sAMAccountName={groupName}))";
                    searcher.SearchScope = SearchScope.Subtree;
                    SearchResult result = searcher.FindOne();
                    if (result != null)
                    {
                        group.Name = result.Properties["name"][0].ToString();
                        group.SamAccountName = result.Properties["sAMAccountName"][0].ToString();
                        group.DistinguishedName = result.Properties["distinguishedName"][0].ToString();
                        group.Description = result.Properties["description"][0].ToString();
                        group.StructuralObjectClass = result.Properties["objectclass"][0].ToString();
                        return group;
                    }
                }
            }
            return null;
        }

        public static bool Exist(string sAMAccountName)
        {
            UserPrincipal user = UserPrincipal.FindByIdentity(principalContext, sAMAccountName);
            if (user != null)
            {
                return true;
            }
            return false;
        }

        public static bool IsMemberOf(string sAMAccountName, string groupName)
        {
            UserPrincipal up = UserPrincipal.FindByIdentity(principalContext, sAMAccountName);
            bool isMember = up.IsMemberOf(principalContext, IdentityType.Name, groupName);
            if (groupName == "Usuarios del dominio" || groupName == "Domain Users")
            {
                isMember = true;
            }
            return isMember;
        }

        public string GetDistinguishedNameFull(string sAMAccountName)
        {
            UserPrincipal user = Get(sAMAccountName);
            if (user != null)
            {
                return user.DistinguishedName;
            }
            return null;
        }

        public string GetUserDn(string userName)
        {

            using (UserPrincipal user = UserPrincipal.FindByIdentity(principalContext, userName))
            {
                if (user != null)
                {
                    DnParser.DN dn = new DnParser.DN(user.DistinguishedName);
                    DnParser.DN parentDn = dn.Parent;
                    return parentDn.ToString();
                }
                throw new Exception("Error al buscar el nombre distinguido");
            }
        }

        public List<Models.Group> getUserGroups(string userName)
        {
            List<Models.Group> groups = new List<Models.Group>();
            UserPrincipal user = UserPrincipal.FindByIdentity(principalContext, IdentityType.SamAccountName, userName);
            foreach (GroupPrincipal group in user.GetGroups())
            {
                groups.Add(new Models.Group
                {
                    Name = group.Name,
                    Description = group.Description,
                    DistinguishedName = group.DistinguishedName,
                    SamAccountName = group.SamAccountName,
                    StructuralObjectClass = group.StructuralObjectClass
                });
            }
            return groups;
        }
    }
}
