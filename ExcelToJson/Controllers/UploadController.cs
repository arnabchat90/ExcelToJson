using Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Mvc;
using Microsoft.Azure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using System.Text.RegularExpressions;
using System.Security;
using Microsoft.SharePoint.Client.UserProfiles;

namespace ExcelToJson.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    public class UploadController : ApiController
    {
        // POST: Upload
        [System.Web.Http.HttpPost, System.Web.Http.Route("api/upload")]
        public async Task<IHttpActionResult> Upload()
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageConnectionString"));
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference("exceltojson");

            container.CreateIfNotExists();

            container.SetPermissions(new BlobContainerPermissions { PublicAccess = BlobContainerPublicAccessType.Blob });

            CloudBlockBlob blockblob = container.GetBlockBlobReference("excelBlob");

            HttpRequestMessage request = this.Request;
            if (!request.Content.IsMimeMultipartContent())
            {
                throw new HttpResponseException(System.Net.HttpStatusCode.UnsupportedMediaType);
            }
            string root = System.Web.HttpContext.Current.Server.MapPath("~/Lib");
            var provider = new MultipartMemoryStreamProvider();
            await Request.Content.ReadAsMultipartAsync(provider);
            foreach (var file in provider.Contents)
            {
                var filename = file.Headers.ContentDisposition.FileName.Trim('\"');
                var buffer = await file.ReadAsByteArrayAsync();

                Stream stream = new MemoryStream(buffer);
                blockblob.UploadFromStream(stream);
                //Do whatever you want with filename and its binaray data.
            }

            return Ok();


            //return await task;
        }

        [System.Web.Http.HttpGet, System.Web.Http.Route("api/getexceljson")]
        public HttpResponseMessage GetExcelJSON()
        {

            #region Using OleDb
            //    var sheetName = "Report Data 1";
            //    var desinationpath = System.Web.HttpContext.Current.Server.MapPath("~/App_Data") + "/audit.json";
            //    var connectionString = String.Format(@"
            //    Provider=Microsoft.ACE.OLEDB.12.0;
            //    Data Source={0};
            //    Extended Properties=""Excel 12.0 Xml;HDR=YES""
            //", filePath);

            //    using (var conn = new OleDbConnection(connectionString))
            //    {
            //        conn.Open();

            //        var cmd = conn.CreateCommand();
            //        cmd.CommandText = String.Format(
            //            @"SELECT * FROM [{0}$]",
            //            sheetName
            //        );


            //        using (var rdr = cmd.ExecuteReader())
            //        {

            //            //LINQ query - when executed will create anonymous objects for each row
            //            var query =
            //                from DbDataRecord row in rdr
            //                select new
            //                {
            //                    name = row[0],
            //                    regno = row[1],
            //                    description = row[2]
            //                };

            //            //Generates JSON from the LINQ query
            //            var json = JsonConvert.SerializeObject(query);

            //            //Write the file to the destination path    
            //            File.WriteAllText(desinationpath, json);

            //            // return json;
            //        }
            //    }
            #endregion

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageConnectionString"));
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            CloudBlobContainer container = blobClient.GetContainerReference("exceltojson");

            container.CreateIfNotExists();

            container.SetPermissions(new BlobContainerPermissions { PublicAccess = BlobContainerPublicAccessType.Blob });
            MemoryStream fileStream = new MemoryStream();
            foreach (IListBlobItem blobItem in container.ListBlobs(null, true))
            {
                CloudBlockBlob blockBlob = (CloudBlockBlob)blobItem;
                //if (blockBlob.Name == "excelBlob")
                //{
                //    blockBlob.DownloadToStream(fileStream);
                //}
                blockBlob.DownloadToStream(fileStream);

            }

            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(fileStream);


            DataSet result = excelReader.AsDataSet();
            var rows = result.Tables[0].Rows;
            List<Dictionary<string, string>> allRows = new List<Dictionary<string, string>>();
            Dictionary<string, string> rowDic;
            var totalRows = rows.Count;
            List<string> customDataCol = new List<string>(20);
            var headRow = rows[2];
            using (ClientContext ctx = new ClientContext("https://wipronvs1.sharepoint.com/sites/siteprov"))
            {
                var passWord = new SecureString();
                foreach (char c in "Sharepoint@2016".ToCharArray()) passWord.AppendChar(c);
                ctx.Credentials = new SharePointOnlineCredentials("surajpadhy@wipronvs1.onmicrosoft.com", passWord);
                var web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                Microsoft.SharePoint.Client.GroupCollection grps = web.SiteGroups;
                ctx.Load(grps, x => x.Include(gr => gr.Title, gr => gr.Id));
                ctx.ExecuteQuery();

                foreach (DataColumn col in result.Tables[0].Columns)
                {
                    customDataCol.Add(headRow[col].ToString().Replace(' ', '_').Replace('(', '_').Replace(')', '_'));
                }
                rows.RemoveAt(0);
                rows.RemoveAt(1);
                //rows.RemoveAt(2);
                foreach (DataRow row in rows)
                {
                    rowDic = new Dictionary<string, string>();
                    var currentEventType = "";
                    for (var i = 0; i < result.Tables[0].Columns.Count; i++)
                    {

                        if (customDataCol[i] == "Occurred__GMT_")
                        {
                            rowDic.Add(customDataCol[i], row[i].ToString().Split('T')[0]);
                        }
                        else if (customDataCol[i] == "Event")
                        {
                            currentEventType = row[i].ToString();
                            rowDic.Add(customDataCol[i], row[i].ToString());
                        }
                        else if (customDataCol[i] == "Event_Data")
                        {
                            var returnString = row[i].ToString();
                            if (currentEventType == "Security Group Member Delete")
                            {
                                Microsoft.SharePoint.Client.Group matchingGroup = null;
                                Regex regex_Group = new Regex("<groupid>(.*?)</groupid>");
                                var v_group = regex_Group.Match(row[i].ToString());
                                var groupID = v_group.Groups[1].ToString();
                                Regex regex_userid = new Regex("<user>(.*?)</user>");
                                var v_userid = regex_userid.Match(row[i].ToString());
                                var usedid = v_userid.Groups[1].ToString();
                                var groupTitle = "";
                                var userDisplayName = "";
                                if (groupID == "") {
                                    groupID = "0";
                                    groupTitle = "-";
                                }
                                else
                                {
                                    matchingGroup = grps.Where(x => x.Id == Convert.ToInt32(groupID)).FirstOrDefault();
                                    groupTitle = matchingGroup.Title;
                                }
                                if (usedid == "")
                                {
                                    usedid = "0";
                                    userDisplayName = "-";
                                }
                                else {
                                    List siteInfoList = ctx.Web.SiteUserInfoList;
                                    var userItem = siteInfoList.GetItemById(Convert.ToInt32(usedid));
                                    ctx.Load(userItem);
                                    ctx.ExecuteQuery();
                                    returnString = "User - " + userItem["Title"].ToString() + " has been deleted from group - " + groupTitle;
                                }
                            }
                            else if (currentEventType == "Security Group Member Add") {
                                Regex regex_Group = new Regex("<groupid>(.*?)</groupid>");
                                var v_group = regex_Group.Match(row[i].ToString());
                                var groupID = v_group.Groups[1].ToString();
                                Regex regex_username = new Regex("<username>(.*?)</username>");
                                var v_username = regex_username.Match(row[i].ToString());
                                var usedName = v_username.Groups[1].ToString();

                                //User targetUser = web.EnsureUser(usedName);
                                //ctx.Load(targetUser);
                                Microsoft.SharePoint.Client.Group matchingGroup = null;
                                if (groupID == "")
                                    groupID = "0";
                                else
                                {
                                    matchingGroup = grps.Where(x => x.Id == Convert.ToInt32(groupID)).FirstOrDefault();
                                }

                                PeopleManager pplMgr = new PeopleManager(ctx);
                                var userDisplayName = "";
                                if (usedName == "")
                                {
                                    userDisplayName = "";
                                }
                                else if (usedName == "c:0(.s|true")
                                {
                                    userDisplayName = "Everyone";
                                }
                                else
                                {
                                    PersonProperties prop = pplMgr.GetPropertiesFor(usedName);
                                    ctx.Load(prop, p => p.AccountName, p => p.DisplayName, p => p.Email);
                                    ctx.ExecuteQuery();
                                    userDisplayName = prop.DisplayName;
                                }
                                
                                    if (userDisplayName == null || userDisplayName == "")
                                        returnString = row[i].ToString();
                                    else
                                        returnString = "User - " + userDisplayName + " was added to the group - " + matchingGroup.Title;
                            }
                           
                            rowDic.Add(customDataCol[i], returnString);


                        }
                        else
                        {
                            rowDic.Add(customDataCol[i], row[i].ToString());
                        }
                    }

                    allRows.Add(rowDic);

                }
            }
           // var dataString = Newtonsoft.Json.JsonConvert.SerializeObject(allRows);
            var jsonString = Newtonsoft.Json.JsonConvert.SerializeObject(new { totalCount = allRows.Count, result = allRows });
            var response = this.Request.CreateResponse(HttpStatusCode.OK);
            response.Content = new StringContent(jsonString, Encoding.UTF8, "application/json");
            return response;


        }
    }
}