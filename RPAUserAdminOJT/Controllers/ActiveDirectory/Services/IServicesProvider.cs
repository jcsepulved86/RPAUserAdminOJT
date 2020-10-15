using System;
using System.Collections.Generic;

namespace RPAUserAdminOJT.Controllers.ActiveDirectory.Services
{
    public interface IServicesProvider
    {
        bool Create(Models.AdUser user);
        void AddUserToGroup(string sAMAccountName, List<Models.Group> groups);
    }
}
