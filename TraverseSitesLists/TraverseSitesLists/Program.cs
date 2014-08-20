using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;

namespace TraverseSitesLists
{
    class Program
    {
        private static bool OnScreen = false;
        private static string SiteURL = "http://tpgnetdev13/";
        private static int allCount = 0;
        private static bool IsSupport = false;
        private static string FileName = "TraverseSiteListPermission.txt";
        private static bool ShowLimitedAccess = false;

        static void Main(string[] args)
        {
            ReadCommandLine(args);
            string msg = "Process Ended at: ";
            Console.WriteLine("Process Started at: " + DateTime.Now);

            if (ShowAndConfirmHelpMessage())
            {

                //Create the outputfile
                CreateToFile();
                StartSiteTraverse();
            }
            else
            {
                msg = "Process terminated by user at: ";
            }

            Console.WriteLine();
            Console.WriteLine(msg + DateTime.Now);
            Console.ReadLine();
        }

        private static void ReadCommandLine(string[] args)
        {
            Arguments commandLine = new Arguments(args);

            if (commandLine["debug"] != null)
            {
                string value = commandLine["debug"].ToString().ToUpper();
                if (value == "YES" || value == "TRUE" || value == "T" || value == "Y")
                    OnScreen = true;
            }

            if (commandLine["support"] != null)
            {
                string value = commandLine["support"].ToString().ToUpper();
                if (value == "YES" || value == "TRUE" || value == "T" || value == "Y")
                    IsSupport = true;
            }

            if (commandLine["showlimitedaccess"] != null)
            {
                string value = commandLine["showlimitedaccess"].ToString().ToUpper();
                if (value == "YES" || value == "TRUE" || value == "T" || value == "Y")
                    ShowLimitedAccess = true;
            }

            if (commandLine["url"] != null)
            {
                SiteURL = commandLine["url"].ToString();
            }

            if (commandLine["filename"] != null)
            {
                FileName = commandLine["filename"].ToString();

                if (FileName.LastIndexOf('.') == -1)
                    FileName += ".txt";
            }
        }

        private static bool ShowAndConfirmHelpMessage()
        {
            bool confirm = false;

            Console.WriteLine(@" ");
            Console.WriteLine(@"Traverse Site List Security.");
            Console.WriteLine(@"Call the program using the following format:");
            Console.WriteLine(@"TraverseSiteList -url=http://tpgnet/ [-filename=Prod_ListPerm.txt] [-showlimitedaccess=No] [-debug=No] [-support=No]");
            Console.WriteLine(@" ");
            Console.WriteLine(@"Parameters: ");
            Console.WriteLine(@"   url               - The starting site you want to traverse");
            Console.WriteLine(@"   filename          - The file name for the output");
            Console.WriteLine(@"   showlimitedaccess - Show Limied Access Roles");
            Console.WriteLine(@"   debug             - [Yes/No] - See the ouput on the screen");
            Console.WriteLine(@"   support           - [Yes/No] - give the support group: texpac\tpgnet_support");
            Console.WriteLine(@"                            the 'Support Level Permission' to the Web and List"); 
            Console.WriteLine(@" ");
            Console.WriteLine(@" ");
            Console.WriteLine(@"Current Parameters: ");
            Console.WriteLine(@"   url               = " + SiteURL );
            Console.WriteLine(@"   filename          = " + FileName);
            Console.WriteLine(@"   showlimitedaccess = " + ShowLimitedAccess.ToString());
            Console.WriteLine(@"   debug             = " + OnScreen.ToString());
            Console.WriteLine(@"   support           = " + IsSupport.ToString() );
            Console.WriteLine(@" ");
            Console.WriteLine(@"Please confirm if the parameters are correct.");
            Console.WriteLine(@"Please type 'Yes' to continue");
            string resp = Console.ReadLine();

            if(resp.ToUpper() == "YES")
                confirm = true;

            return confirm;
        }

        private static void StartSiteTraverse()
        {
            int index = 0;

            string siteName = string.Empty;

            using (SPSite site = new SPSite(SiteURL))
            {
                SPWebCollection webs = site.OpenWeb().Webs; //.AllWebs;
                for (int i = 0; i < webs.Count; i++)
                {
                    try
                    {
                        using (SPWeb web = webs[i])
                        {
                            siteName = web.Title;

                            if (web.HasUniqueRoleAssignments)
                            {
                                WriteToFile("Site " + index++.ToString() + ": " + web.Url);

                                SPRoleAssignmentCollection oRoleAssignments = web.RoleAssignments;
                                PrintRoleAssignments(oRoleAssignments);

                                if (IsSupport)
                                {
                                    GiveWebSuportPermission(web);
                                }
                            }

                            WriteToFile("");

                            //SPListCollection lists = web.Lists;

                            StartListTraverse(web);
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteToFile("");
                        WriteToFile("~~~ ERROR Site: " + siteName + " - " + ex.Message);
                        WriteToFile("");
                    }
                }

                WriteToFile("Total sites: " + webs.Count.ToString());
                WriteToFile("Grand Total sites and Lists: " + allCount.ToString());
            }
        }


        private static void StartListTraverse(SPWeb web)
        {
            if (web == null)
                return;

            string listName = string.Empty;

            try
            {
                SPListCollection lists = web.Lists;

                for (int j = 0; j < lists.Count; j++)
                {
                    listName = lists[j].Title;

                    if (!lists[j].Hidden)
                    {
                        if (lists[j].HasUniqueRoleAssignments)
                        {
                            allCount++;
                            WriteToFile("\t List: " + web.Url + lists[j].RootFolder.Url + " - " + lists[j].Title);

                            SPRoleAssignmentCollection oRoleAssignments = lists[j].RoleAssignments;
                            PrintRoleAssignments(oRoleAssignments);

                            WriteToFile("");

                            if (IsSupport)
                            {
                                GiveListSuportPermission(web, lists[j]);
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile("");
                WriteToFile("~~~ ERROR - List: "  + listName + " - " + ex.Message);
                WriteToFile("");
            }
        }

        private static void PrintRoleAssignments(SPRoleAssignmentCollection oRoleAssignments)
        {
            string message = string.Empty;
            string roles = string.Empty;

            foreach (SPRoleAssignment oRoleAssignment in oRoleAssignments)
            {

                SPPrincipal oPrincipal = oRoleAssignment.Member;
                try
                {
                    if (oPrincipal is SPUser)
                    {
                        // Retrieve users having explicit permissions on the list
                        SPUser oRoleUser = (SPUser)oPrincipal;

                        if (oRoleUser.IsDomainGroup)
                            message = "\t\t Domain: ";
                        else
                            message = "\t\t User: ";

                        message += oRoleUser.Name;

                        //Console.WriteLine("        User: " + oRoleUser.Name);
                        //WriteToFile("        User: " + oRoleUser.Name);

                        roles = string.Empty;

                        foreach (SPRole role in oRoleUser.Roles)
                        {
                            if (!ShowLimitedAccess)
                            {
                                if (role.Name.ToUpper() != "LIMITED ACCESS")
                                {
                                    roles += role.Name + ", ";
                                }
                            }
                            else
                            {
                                roles += role.Name + ", ";
                                //Console.WriteLine("          User Roles: " + role.Name);
                                //WriteToFile("          User Roles: " + role.Name);
                            }
                        }

                        if (!string.IsNullOrEmpty(roles))
                        {
                            roles = roles.Substring(0, roles.Length - 2);

                            //Console.WriteLine(message + "( " + roles + " )");
                            WriteToFile(message + "( " + roles + " )");
                        }


                    }
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("** Error (User): " + ex.Message);
                    WriteToFile("** Error (User): " + ex.Message);
                }
                try
                {
                    if (oPrincipal is SPGroup)
                    {
                        // Retrieve user groups having permissions on the list
                        SPGroup oRoleGroup = (SPGroup)oPrincipal;

                        string strGroupName = oRoleGroup.Name;
                        // Add code here to retrieve Users inside this User-Group
                        //Console.WriteLine("        Group: " + strGroupName);
                        //WriteToFile("        Group: " + strGroupName);

                        roles = string.Empty;
                        foreach (SPRole role in oRoleGroup.Roles)
                        {
                            if (!ShowLimitedAccess)
                            {
                                if (role.Name.ToUpper() != "LIMITED ACCESS")
                                {
                                    roles += role.Name + ", ";
                                }
                            }
                            else
                            {
                                //Console.WriteLine("          Group Roles: " + role.Name);
                                //WriteToFile("          Group Roles: " + role.Name);
                                roles += role.Name + ", ";
                            }
                        }

                        if (!string.IsNullOrEmpty(roles))
                        {
                            roles = roles.Substring(0, roles.Length - 2);


                            //Console.WriteLine("        Group: " + strGroupName + "( " + roles + " )");
                            WriteToFile("\t\t Group: " + strGroupName + "( " + roles + " )");

                            if (oRoleGroup.Users.Count > 0)
                            {

                                foreach (SPUser user in oRoleGroup.Users)
                                {
                                    if (user.IsDomainGroup)
                                        message = "\t\t\t Domain: ";
                                    else
                                        message = "\t\t\t User: ";

                                    //Console.WriteLine(message + user.Name);
                                    WriteToFile(message + user.Name);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("** Error (Group): " + ex.Message);
                    WriteToFile("** Error (Group): " + ex.Message);
                }
            }
        }

        private static void GiveWebSuportPermission(SPWeb web)
        {
            try
            {
                web.AllowUnsafeUpdates = true;
                web.Update();

                SPUser user = web.EnsureUser(@"texpac\tpgnet_support");

                SPRoleAssignment roleAssign = new SPRoleAssignment(user);
                SPRoleDefinition roleDef = web.RoleDefinitions["TPGNet Support Permission Level"];
                roleAssign.RoleDefinitionBindings.Add(roleDef);
                web.RoleAssignments.Add(roleAssign);
                web.Update();
            }
            catch (Exception ex)
            {
                WriteToFile("Error: SupportPermission - " + ex.Message);
            }
        }

        private static void GiveListSuportPermission(SPWeb web, SPList list)
        {
            try
            {
                web.AllowUnsafeUpdates = true;
                web.Update();

                SPUser user = web.EnsureUser(@"texpac\tpgnet_support");

                SPRoleAssignment roleAssign = new SPRoleAssignment(user);
                SPRoleDefinition roleDef = web.RoleDefinitions["TPGNet Support Permission Level"];
                roleAssign.RoleDefinitionBindings.Add(roleDef);
                list.RoleAssignments.Add(roleAssign);
                list.Update();
                web.Update();
            }
            catch (Exception ex)
            {
                WriteToFile("Error: SupportPermission - " + ex.Message);
            }
        }

        private static void WriteToFile(string text)
        {
            using (StreamWriter writer = new StreamWriter(FileName, true))
            {
                writer.WriteLine(text);
            }

            if (OnScreen)
                Console.WriteLine(text);
        }

        private static void CreateToFile()
        {
            StreamWriter writer = new StreamWriter(FileName, false);
            //using (StreamWriter writer = new StreamWriter("debug.txt", false))
            //{
            //    writer.WriteLine("");
            //}
        }
    }
}
