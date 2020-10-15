﻿using System;
using System.Collections.Generic;
using System.Configuration;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.IO;

namespace RPAUserAdminOJT.Controllers.ActiveDirectory.Services
{
    public class ServicesProvider : IServicesProvider
    {
        public bool result;
        public static PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, Models.GlobalVar.Domain, Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);
        Query.QueryProvider MainQueryProvider;
        private string password;

        public static string rootblackList = ConfigurationManager.AppSettings["rootblackList"];

        public ServicesProvider(PrincipalContext principalContext, string password = null)
        {
            this.password = password;
            principalContext = principalContext ?? throw new ArgumentNullException(nameof(principalContext));
            this.MainQueryProvider = new Query.QueryProvider(principalContext);
        }

        /// <summary>
        /// Crea un nuevo usuario en el AD
        /// </summary>
        /// <param name="user">Objecto de tipo ADuser</param>
        public bool Create(Models.AdUser user)
        {
            try
            {   // Valida que el usuario no exista
                if (!Query.QueryProvider.Exist(user.SamAccountName))
                {

                    string oGUID = string.Empty;
                    string connectionPrefix = $"LDAP://{user.DistinguishedName}";
                    DirectoryEntry dirEntry = new DirectoryEntry(connectionPrefix, principalContext.UserName, password);
                    DirectoryEntry newUser;
                    // Valida que el Nombre para mostrar no exista
                    string tempDn = $"CN={user.GetFullName()},{user.DistinguishedName}";
                    var u = Query.QueryProvider.Get(tempDn);
                    if (u != null)
                    {
                        newUser = dirEntry.Children.Add("CN=" + user.GetFullName() + ".", "user");
                        newUser.Properties["cn"].Value = user.GetFullName() + ".";
                        newUser.Properties["displayName"].Value = user.GetFullName() + ".";
                    }
                    else
                    {
                        newUser = dirEntry.Children.Add("CN=" + user.GetFullName(), "user");
                        newUser.Properties["cn"].Value = user.GetFullName();
                        newUser.Properties["displayName"].Value = user.GetFullName();
                    }
                    newUser.Properties["samAccountName"].Value = user.SamAccountName;
                    newUser.Properties["Description"].Value = user.Description;
                    newUser.Properties["userPrincipalName"].Value = user.GetUserPrincipalName();
                    newUser.Properties["sn"].Value = user.Surname();
                    //newUser.Properties["mail"].Value = "";
                    newUser.Properties["givenName"].Value = user.GetName();
                    newUser.Properties["postOfficeBox"].Value = user.PostOfficeBox;
                    newUser.Properties["l"].Value = user.City;
                    newUser.Properties["st"].Value = user.State;
                    newUser.Properties["co"].Value = user.Country;
                    newUser.Properties["department"].Value = user.Department;
                    newUser.Properties["company"].Value = user.Company;
                    newUser.Properties["wWWHomePage"].Value = user.GetHomePage();
                    newUser.CommitChanges();
                    oGUID = newUser.Guid.ToString();
                    newUser.Invoke("SetPassword", new object[] { user.Password });
                    newUser.CommitChanges();
                    dirEntry.Close();
                    newUser.Close();

                    return result = true;

                }
                else
                {
                    return result = false;
                }
            }
            catch (Exception ex)
            {
                return result = false;
                throw ex;
            }
        }

        public bool Create(Models.AdUser user, bool enabled)
        {
            string CrtMessage = string.Empty;

            try
            {   // Valida que el usuario no exista
                if (!Query.QueryProvider.Exist(user.SamAccountName))
                {

                    string oGUID = string.Empty;
                    string connectionPrefix = $"LDAP://{user.DistinguishedName}";
                    DirectoryEntry dirEntry = new DirectoryEntry(connectionPrefix, principalContext.UserName, password);
                    DirectoryEntry newUser;
                    // Valida que el Nombre para mostrar no exista
                    string tempDn = $"CN={user.GetFullName()},{user.DistinguishedName}";
                    var u = Query.QueryProvider.Get(tempDn);
                    if (u != null)
                    {
                        newUser = dirEntry.Children.Add("CN=" + user.GetFullName() + ".", "user");
                        newUser.Properties["cn"].Value = user.GetFullName() + ".";
                        newUser.Properties["displayName"].Value = user.GetFullName() + ".";
                    }
                    else
                    {
                        newUser = dirEntry.Children.Add("CN=" + user.GetFullName(), "user");
                        newUser.Properties["cn"].Value = user.GetFullName();
                        newUser.Properties["displayName"].Value = user.GetFullName();
                    }
                    newUser.Properties["samAccountName"].Value = user.SamAccountName;
                    newUser.Properties["Description"].Value = user.Description;
                    newUser.Properties["userPrincipalName"].Value = user.GetUserPrincipalName();
                    newUser.Properties["sn"].Value = user.Surname();
                    //newUser.Properties["mail"].Value = "";
                    newUser.Properties["givenName"].Value = user.GetName();
                    newUser.Properties["postOfficeBox"].Value = user.PostOfficeBox;
                    newUser.Properties["l"].Value = user.City;
                    newUser.Properties["st"].Value = user.State;
                    newUser.Properties["co"].Value = user.Country;
                    newUser.Properties["department"].Value = user.Department;
                    newUser.Properties["company"].Value = user.Company;
                    newUser.Properties["wWWHomePage"].Value = user.GetHomePage();

                    newUser.CommitChanges();
                    oGUID = newUser.Guid.ToString();
                    newUser.Invoke("SetPassword", new object[] { user.Password });

                    if (enabled)
                    {
                        EnableAccount(newUser);  // Habilita la cuenta
                    }
                    newUser.CommitChanges();
                    dirEntry.Close();
                    newUser.Close();

                    Models.GlobalVar.accessEntires = true;
                    return result = true;

                }
                else
                {
                    Models.GlobalVar.accessEntires = true;
                    return result = false;
                }
            }
            catch (Exception e)
            {
                CrtMessage = e.Message.ToString();

                if (CrtMessage.Contains("Acceso denegado"))
                {
                    Models.GlobalVar.accessEntires = false;
                }
                else if (CrtMessage.Contains("Access denied."))
                {
                    Models.GlobalVar.accessEntires = false;
                }

                return result = false;
                throw e;

            }
        }

        /// <summary>
        /// Agrega un usuario a un grupo
        /// </summary>
        /// <param name="sAMAccountName">Usuario de red</param>
        /// <param name="groups">Listado de grupos</param>
        public void AddUserToGroup(string sAMAccountName, List<Models.Group> groups)
        {
            string groupAcces = string.Empty;
            string accessMessage = string.Empty;
            bool varaxu = false;


            try
            {
                foreach (Models.Group group in groups)
                {
                    if (!Query.QueryProvider.IsMemberOf(sAMAccountName, group.Name))
                    {

                        GroupPrincipal gp = GroupPrincipal.FindByIdentity(principalContext, group.Name);

                        string result = string.Empty;
                        string[] filelines = File.ReadAllLines(rootblackList);

                        for (int a = 0; a < filelines.Length; a++)
                        {
                            string compare = filelines[a];

                            if (gp.Name == compare)
                            {
                                varaxu = true;
                                break;
                            }

                        }

                        if (varaxu == true)
                        {
                            continue;
                        }
                        else
                        {
                            try
                            {
                                gp.Members.Add(principalContext, IdentityType.SamAccountName, sAMAccountName);
                                gp.Save();
                            }
                            catch (Exception e)
                            {
                                accessMessage = e.Message.ToString();

                                if (accessMessage.Contains("Acceso denegado"))
                                {
                                    //Controllers.LogFiles.LogApplication.LogWrite("User AddGrp ==> " + "Usuario: " + sAMAccountName + ", esta directiva no se puede asignar: " + gp.Name);
                                    continue;
                                }
                                else if (accessMessage.Contains("Access denied."))
                                {
                                    //Controllers.LogFiles.LogApplication.LogWrite("User AddGrp ==> " + "Usuario: " + sAMAccountName + ", esta directiva no se puede asignar: " + gp.Name);
                                    continue;
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }


                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary> 
        /// Habilita la cuenta de usuario
        /// </summary>
        public static void EnableAccount(DirectoryEntry de)
        {
            //UF_ACCOUNTDISABLE 0x0002
            int val = (int)de.Properties["userAccountControl"].Value;
            de.Properties["userAccountControl"].Value = val & ~0x0002;
            de.CommitChanges();
        }


        public static void User_Add_Group(string users, string baseUser)
        {

            PrincipalContext principalContext = new PrincipalContext(ContextType.Domain, Models.GlobalVar.Domain, Models.GlobalVar.UserAdm, Models.GlobalVar.Passwrd);
            ServicesProvider servicesProvider = new ServicesProvider(principalContext);
            UserPrincipal user = Query.QueryProvider.Get(users);
            List<Models.Group> lstGroups = Query.QueryProvider.GetGroups(baseUser);
            servicesProvider.AddUserToGroup(users, lstGroups);

            user.Dispose();
            principalContext.Dispose();

        }


        public static void PasswordNeverExpires(string users)
        {
            UserPrincipal user = Query.QueryProvider.Get(users);
            user.Enabled = true;
            user.ExpirePasswordNow();
            user.PasswordNeverExpires = true;
            user.Save();

            user.Dispose();
            principalContext.Dispose();
        }

        public static void DisableAccount(string users)
        {
            UserPrincipal user = Query.QueryProvider.Get(users);
            user.Enabled = false;
            user.Save();

            user.Dispose();
            principalContext.Dispose();
        }


    }
}