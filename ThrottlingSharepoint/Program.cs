using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Newtonsoft.Json;
using System;
using System.Net;
using System.ServiceModel.Description;
using Microsoft.SharePoint.Client;
using System.Security;
using System.IO;

namespace ThrottlingSharepoint
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            ReadConfigFile();
            sp_fetchRecords();
        }

        public static void sp_fetchRecords()
        {
            var watch = new System.Diagnostics.Stopwatch();
            watch.Start();
            ClientContext sharepointContext = SharepointConnection();
            if(sharepointContext == null)
            {
                return;
            }
            Console.WriteLine("Sharepoint Authenticated\n");
            var clientCredentials = CRMConnection();

            string text = System.IO.File.ReadAllText("C:/Users/utgoswam/source/repos/ThrottlingSharepoint/ThrottlingSharepoint/fileNumber.txt");
            int fileNumber = Convert.ToInt32(text);

            // create files to log successful and failed record movements
            successFile = System.IO.File.Create(string.Format("C:/Users/utgoswam/source/repos/ThrottlingSharepoint/ThrottlingSharepoint/success{0}.xml", fileNumber));
            successWriter = new System.Xml.Serialization.XmlSerializer(typeof(Success));
            failureFile = System.IO.File.Create(string.Format("C:/Users/utgoswam/source/repos/ThrottlingSharepoint/ThrottlingSharepoint/failures{0}.xml", fileNumber));
            failureWriter = new System.Xml.Serialization.XmlSerializer(typeof(Failure));
            workingFile = System.IO.File.Create(string.Format("C:/Users/utgoswam/source/repos/ThrottlingSharepoint/ThrottlingSharepoint/working{0}.xml", fileNumber));
            workingWriter = new System.Xml.Serialization.XmlSerializer(typeof(Success));

            System.IO.File.WriteAllText("C:/Users/utgoswam/source/repos/ThrottlingSharepoint/ThrottlingSharepoint/fileNumber.txt", Convert.ToString(fileNumber + 1));

            // For Dynamics 365 Customer Engagement V9.X, set Security Protocol as TLS12
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            // Connect to the Organization service. 
            // The using statement assures that the service proxy will be properly disposed.
            using (var _serviceProxy = new OrganizationServiceProxy(new Uri(crmurl + "/XRMServices/2011/Organization.svc"), null, clientCredentials, null))
            {
                QueryExpression query = new QueryExpression
                {
                    EntityName = "sharepointdocumentlocation",
                    ColumnSet = new ColumnSet("sharepointdocumentlocationid", "regardingobjectid", "parentsiteorlocation", "relativeurl"),
                    PageInfo = new PagingInfo()
                };
                query.PageInfo.Count = 5000;
                query.PageInfo.PageNumber = 1;
                query.PageInfo.ReturnTotalRecordCount = true;

                EntityCollection sharepointDocumentLocations;

                do
                {
                    sharepointDocumentLocations = _serviceProxy.RetrieveMultiple(query);
                    updateDocumentLocationsAndMoveFiles(sharepointDocumentLocations, _serviceProxy, sharepointContext);

                    query.PageInfo.PageNumber += 1;
                    query.PageInfo.PagingCookie = sharepointDocumentLocations.PagingCookie;
                } while (sharepointDocumentLocations.MoreRecords);
            }
            watch.Stop();
            Console.WriteLine($"Execution Time: {watch.ElapsedMilliseconds} ms");
            Console.WriteLine("The script has completed successfully. Press any key to end...");
            Console.ReadKey();
            successFile.Close();
            workingFile.Close();
            failureFile.Close();
        }

        private static void updateDocumentLocationsAndMoveFiles(EntityCollection sharepointDocumentLocations, OrganizationServiceProxy serviceProxy, ClientContext sharepointContext)
        {
            foreach (Entity sharepointDocumentLocation in sharepointDocumentLocations.Entities)
            {
                //check for SDL not being eccentric and then only proceed
                if (sharepointDocumentLocation.Contains("regardingobjectid"))
                {
                    EntityReference entity = (EntityReference)sharepointDocumentLocation["regardingobjectid"];
                    if (entity.LogicalName == entityLogicalName && sharepointDocumentLocation.Contains("parentsiteorlocation"))
                    {
                        EntityReference parentSiteOrLocation = (EntityReference)sharepointDocumentLocation["parentsiteorlocation"];
                        Entity parentDL = serviceProxy.Retrieve(parentSiteOrLocation.LogicalName, parentSiteOrLocation.Id, new ColumnSet("relativeurl", "parentsiteorlocation"));
                        if (parentDL.Contains("relativeurl") && (string)parentDL["relativeurl"] == entityLogicalName)
                        {
                            if (count % numberOfFilesPerDL == 0)
                            {
                                EntityReference parentDefaultSite = (EntityReference)parentDL["parentsiteorlocation"];
                                Entity newSharepointDL = new Entity("sharepointdocumentlocation");
                                newRelativeUrl = suffixForNewDL + '_' + documentLibraryCount;
                                documentLibraryCount++;
                                newSharepointDL["name"] = newRelativeUrl;
                                newSharepointDL["parentsiteorlocation"] = parentDefaultSite;
                                newSharepointDL["relativeurl"] = newRelativeUrl;
                                newDocumentLocationID = serviceProxy.Create(newSharepointDL);
                            }
                            sharepointDocumentLocation["parentsiteorlocation"] = new EntityReference("sharepointdocumentlocation", newDocumentLocationID);
                            try
                            {
                                Success currentEntity = new Success() { EntityId = entity.Id, SDLId = (Guid)sharepointDocumentLocation.Attributes["sharepointdocumentlocationid"], LogicalName = entity.LogicalName, Name = entity.Name };
                                workingWriter.Serialize(workingFile, currentEntity);
                                Console.WriteLine("Trying to copy the files and update the sharepointdocumentlocation for {0} with ID {1}", entityLogicalName, entity.Id);
                                Folder sourceFolder = CopyFiles(sharepointContext, newRelativeUrl, sharepointDocumentLocation);
                                serviceProxy.Update(sharepointDocumentLocation);
                                // delete only if customer has agreed to delete files in the config file
                                if (shouldDeleteFiles)
                                {
                                    sourceFolder.DeleteObject();
                                    sharepointContext.ExecuteQuery();
                                }
                                successWriter.Serialize(successFile, currentEntity);
                            }
                            catch (Exception ex)
                            {
                                Failure failureLog = new Failure() { entityRef = entity, SDLId = (Guid)sharepointDocumentLocation.Attributes["sharepointdocumentlocationid"], errorMessage = ex.Message };
                                failureWriter.Serialize(failureFile, failureLog);
                                Console.WriteLine("Failed to copy files or update the sharepointdocumentlocation for {0} with ID {1}", entityLogicalName, entity.Id);
                                break;
                            }
                            // increasing the count irrespective of failure because files might have been moved
                            count++;
                            if (count == limitRecords) break;
                        }
                    }
                }
            }
        }

        private static Folder CopyFiles(ClientContext context, string newRelativeUrl, Entity sharepointdocumentlocation)
        {
            var sourceFolder = context.Web.GetFolderByServerRelativeUrl(GetRelativeURLForDL() + sharepointdocumentlocation["relativeurl"]);
            CopyFilesToAnotherDL(sourceFolder, GetRelativeURLForNewDL() + newRelativeUrl + "/" + sharepointdocumentlocation["relativeurl"], newRelativeUrl, context);
            return sourceFolder;
        }

        private static ClientCredentials CRMConnection()
        {
            ClientCredentials clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = crmusername;
            Console.WriteLine("Please enter the password for CRM user {0}", crmusername);
            clientCredentials.UserName.Password = GetPassword();
            Console.WriteLine("\nTrying to authenticate CRM credentials and fetch data...\n");
            return clientCredentials;
        }

        private static ClientContext SharepointConnection()
        {
            Console.WriteLine("Please enter the password for Sharepoint user ID {0}", spusername);
            string password = GetPassword();
            Console.WriteLine("\nTrying to authenticate sharepoint credentials...");
            return SharePointAuth(spusername, password, sharepointurl);
        }

        private static string GetPassword()
        {
            string password = "";
            ConsoleKeyInfo nextKey = Console.ReadKey(true);
            while (nextKey.Key != ConsoleKey.Enter)
            {
                if (nextKey.Key == ConsoleKey.Backspace)
                {
                    if (password.Length > 0)
                    {
                        password = password.Substring(0, password.Length - 1);
                        Console.Write(nextKey.KeyChar);
                        Console.Write(" ");
                        Console.Write(nextKey.KeyChar);
                    }
                }
                else
                {
                    password += nextKey.KeyChar;
                    Console.Write("*");
                }
                nextKey = Console.ReadKey(true);
            }
            return password;
        }

        private static ClientContext SharePointAuth(String username, String pwd, string siteURL)
        {
            ClientContext context = new ClientContext(siteURL);
            Web web = context.Web;
            SecureString passWord = new SecureString();
            foreach (char c in pwd.ToCharArray()) passWord.AppendChar(c);
            context.Credentials = new SharePointOnlineCredentials(username, passWord);
            try
            {
                context.Load(web);
                context.ExecuteQuery();
                return context;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return null;
            }
        }

        private static void CopyFilesToAnotherDL(this Folder folder, string destFolderUrl, string documentLibraryName, ClientContext context)
        {
            var ctx = (ClientContext)folder.Context;
            if (!ctx.Web.IsPropertyAvailable("ServerRelativeUrl"))
            {
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
            }
            ctx.Load(folder, f => f.Files, f => f.ServerRelativeUrl, f => f.Folders);
            ctx.ExecuteQuery();

            EnsureDL(context, documentLibraryName);

            //ctx.Load(folder.Folders);
            //ctx.ExecuteQuery();
            EnsureFolder(ctx.Web.RootFolder, destFolderUrl.Replace(ctx.Web.ServerRelativeUrl, string.Empty));
            foreach (var file in folder.Files)
            {
                var targetFileUrl = file.ServerRelativeUrl.Replace(folder.ServerRelativeUrl, destFolderUrl);
                file.CopyTo(targetFileUrl, true);
            }
            ctx.ExecuteQuery();

            foreach (var subFolder in folder.Folders)
            {
                var targetFolderUrl = subFolder.ServerRelativeUrl.Replace(folder.ServerRelativeUrl, destFolderUrl);
                CopyFilesToAnotherDL(subFolder, targetFolderUrl, documentLibraryName, context);
                if (subFolder.ProgID == "OneNote.Notebook")
                {
                    List srcList = ctx.Web.Lists.GetByTitle(documentLibraryName);
                    var qry = CamlQuery.CreateAllItemsQuery();
                    qry.FolderServerRelativeUrl = targetFolderUrl.Replace(subFolder.Name, string.Empty);
                    var srcItems = srcList.GetItems(qry);
                    ctx.Load(srcItems, icol => icol.Include(i => i.FileSystemObjectType, i => i["FileRef"], i => i.File));
                    ctx.ExecuteQuery();
                    foreach (ListItem item in srcItems)
                    {
                        if (item.FieldValues["FileRef"].ToString() == targetFolderUrl)
                        {
                            item["HTML_x0020_File_x0020_Type"] = "OneNote.Notebook";
                            item.SystemUpdate();
                            ctx.ExecuteQuery();
                            break;
                        }
                    }
                }
            }
        }

        private static void EnsureDL(ClientContext context, string documentLibraryName)
        {
            ListCreationInformation documentLibrary = new ListCreationInformation();
            documentLibrary.Title = documentLibraryName;
            documentLibrary.Description = descriptionOfNewDL;
            documentLibrary.TemplateType = (int)ListTemplateType.DocumentLibrary;
            ListCollection lists = context.Web.Lists;
            context.Load(lists);
            context.ExecuteQuery();
            foreach (var listItem in lists)
            {
                if (listItem.Title == documentLibraryName)
                {
                    return;
                }
            }
            context.Web.Lists.Add(documentLibrary);
            context.ExecuteQuery();
        }

        private static Folder EnsureFolder(Folder parentFolder, string folderUrl)
        {
            var ctx = parentFolder.Context;
            var folderNames = folderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var folderName = folderNames[0];
            var folder = parentFolder.Folders.Add(folderName);
            ctx.Load(folder);
            ctx.ExecuteQuery();

            if (folderNames.Length > 1)
            {
                var subFolderUrl = string.Join("/", folderNames, 1, folderNames.Length - 1);
                return EnsureFolder(folder, subFolderUrl);
            }
            return folder;
        }

        private static string GetRelativeURLForDL()
        {
            return '/' + subsiteRelativeUrl.TrimEnd('/').TrimStart('/') + '/' + entityLogicalName.TrimStart('/').TrimEnd('/') + '/';
        }

        private static string GetRelativeURLForNewDL()
        {
            return '/' + subsiteRelativeUrl.TrimEnd('/').TrimStart('/') + '/';
        }

        private static void ReadConfigFile()
        {
            using (StreamReader r = new StreamReader("C:/Users/utgoswam/source/repos/ThrottlingSharepoint/ThrottlingSharepoint/config.json"))
            {
                string json = r.ReadToEnd();
                Config items = JsonConvert.DeserializeObject<Config>(json);
                crmurl = items.crmUrl;
                crmusername = items.crmUserName;
                sharepointurl = items.sharepointUrl + '/' + (items.isThereAnySubsite.ToLower() == "yes" ? items.subsiteUrl : "");
                spusername = items.sharepointUserEmailID;
                entityLogicalName = items.entityLogicalName;
                numberOfFilesPerDL = int.Parse(items.numberOfFoldersPerDL);
                suffixForNewDL = items.newDLPrefix;
                documentLibraryCount = int.Parse(items.newDLSuffixNumber);
                subsiteRelativeUrl = (items.isThereAnySubsite.ToLower() == "yes" ? items.subsiteUrl : "");
                shouldDeleteFiles = (items.deleteFilesAfterCopying.ToLower() == "yes" ? true : false);
                descriptionOfNewDL = items.descriptionOfNewDL;
                limitRecords = int.Parse(items.maxRecordsToBeMoved);
            }
        }

        static string crmurl;
        static string crmusername;
        static string sharepointurl;
        static string spusername;
        static string entityLogicalName;
        static string subsiteRelativeUrl = "";
        static int numberOfFilesPerDL = 3000;
        static int documentLibraryCount = 0;
        static string suffixForNewDL = "";
        static bool shouldDeleteFiles = false;
        static string descriptionOfNewDL;
        static int limitRecords = Int32.MaxValue;

        static int count = 0;
        static Guid newDocumentLocationID = Guid.Empty;
        static string newRelativeUrl = "";

        static System.Xml.Serialization.XmlSerializer successWriter;
        static System.Xml.Serialization.XmlSerializer failureWriter;
        static System.Xml.Serialization.XmlSerializer workingWriter;
        static FileStream successFile;
        static FileStream failureFile;
        static FileStream workingFile;
    }
}
