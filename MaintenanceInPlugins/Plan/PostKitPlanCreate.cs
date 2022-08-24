namespace MaintenanceIn.MaintenanceInPlugins
{
    using System;
    using System.Data;
    using Microsoft.Xrm.Sdk;
    using System.Data.SqlClient;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Crm.Sdk.Messages;
    using System.Data.Sql;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using System.Threading.Tasks;

    public class PostKitPlanCreate : PluginBase
    {
        public PostKitPlanCreate(string unsecure, string secure) : base(typeof(PostKitPlanCreate))
        { }

        protected string result = string.Empty;
        /// <summary>
        /// backup path
        /// </summary>
        string pathBackup = string.Empty;
        /// <summary>
        /// path for copy database
        /// </summary>
        string shareFolder = string.Empty;
        protected override void ExecuteCrmPlugin(LocalPluginContext localContext)
        {
            var position = 1d;
            if (localContext == null)
            {
                throw new InvalidPluginExecutionException("localContext");
            }

            position = 2;
            var context = localContext.PluginExecutionContext;
            var service = localContext.OrganizationService;
            //setting variables
            var sqlUser = string.Empty;
            var sqlPassword = string.Empty;
            var serverName = string.Empty;
            var databaseName = context.OrganizationName + "_MSCRM";

            //database file name 
            var name = string.Empty;

            var domainUsername = string.Empty;
            var domainPassword = string.Empty;
            var domain = string.Empty;
            var copytoshare = false;
            var desfolder = string.Empty;

            var retCurrent = new EntityReference("ymb_plan");
            var current = new Entity(context.PrimaryEntityName);

            try
            {
                if (context.Depth < 2)
                {
                    if (context.MessageName.Equals("Create", StringComparison.InvariantCultureIgnoreCase))
                    {
                        position = 3;
                        current = (context.InputParameters.Contains("Target") &&
                                    context.InputParameters["Target"] != null) ?
                                            (Entity)context.InputParameters["Target"] : null;

                        //default path
                        pathBackup = @"C:\DynamicsBackup";
                        //default instance
                        serverName = ".";

                        //----------------------------------------------------------------------------
                        #region Read Settings
                        var FE = new FilterExpression(LogicalOperator.Or);
                        //sql sa user
                        FE.AddCondition("ymb_key", ConditionOperator.Equal, "SQLUSER");
                        //sql sa password
                        FE.AddCondition("ymb_key", ConditionOperator.Equal, "SQLPASSWORD");
                        //instance name
                        FE.AddCondition("ymb_key", ConditionOperator.Equal, "SERVERNAME");
                        //backup path
                        FE.AddCondition("ymb_key", ConditionOperator.Equal, "BACKUPPATH");

                        //these setting are for copy file db to a sharefolder
                        //domain admin
                        FE.AddCondition("ymb_key", ConditionOperator.Equal, "DOMAINUSER");
                        //domain admin password
                        FE.AddCondition("ymb_key", ConditionOperator.Equal, "DOMAINPASSWORD");
                        //domain name
                        FE.AddCondition("ymb_key", ConditionOperator.Equal, "DOMAIN");
                        //folder name for copy backup 
                        FE.AddCondition("ymb_key", ConditionOperator.Equal, "SHAREFOLDER");

                        //*************************************************************************
                        //
                        // you can save all above setting in entity for example ymb_basesetting.
                        //
                        //*************************************************************************
                        var QE = new QueryExpression
                        {
                            EntityName = "ymb_basesetting",
                            ColumnSet = new ColumnSet("ymb_key", "ymb_value"),
                            Criteria = FE
                        };

                        position = 3.001;
                        var retSetting = service.RetrieveMultiple(QE);

                        if (retSetting.Entities.Count > 0)
                        {
                            position = 3.01;
                            foreach (var item in retSetting.Entities)
                            {
                                if (item.Contains("ymb_key"))
                                {
                                    switch (item["ymb_key"].ToString())
                                    {
                                        case "SQLUSER":
                                            sqlUser = item.Contains("ymb_value") ? item["ymb_value"].ToString() : "";
                                            break;
                                        case "SQLPASSWORD":
                                            sqlPassword = item.Contains("ymb_value") ? item["ymb_value"].ToString() : "";
                                            break;
                                        case "SERVERNAME":
                                            serverName = item.Contains("ymb_value") ? item["ymb_value"].ToString() : "";
                                            break;
                                        case "BACKUPPATH":
                                            pathBackup = item.Contains("ymb_value") ? item["ymb_value"].ToString() : "";
                                            break;
                                        case "DOMAINUSER":
                                            domainUsername = item.Contains("ymb_value") ? item["ymb_value"].ToString() : "";
                                            break;
                                        case "DOMAINPASSWORD":
                                            domainPassword = item.Contains("ymb_value") ? item["ymb_value"].ToString() : "";
                                            break;
                                        case "DOMAIN":
                                            domain = item.Contains("ymb_value") ? item["ymb_value"].ToString() : "";
                                            break;
                                        case "SHAREFOLDER":
                                            shareFolder = item.Contains("ymb_value") ? item["ymb_value"].ToString() : "";
                                            break;
                                    }
                                }
                            }
                        }
                        #endregion
                        //
                        if (current.Contains("ymb_backup") && (bool)current["ymb_backup"])
                        {
                            position = 3.02;
                            #region Check for minimum drive space (optional)

                            DriveInfo[] allDrives = DriveInfo.GetDrives();
                            foreach (DriveInfo drive in allDrives)
                            {
                                if (drive.Name.StartsWith(pathBackup.Substring(0, 1)))
                                {
                                    //check for free space in drive (optional)
                                    if (drive.TotalFreeSpace < 21474836480)
                                    {
                                        //show the result to the user in a field
                                        current.Attributes.Add("ymb_backupresult",
                                            string.Format("Drive {0} Total size: {1} MB Total free space: {2} MB",
                                            drive.Name,
                                            ConvertToMegaByte(drive.TotalSize),
                                            ConvertToMegaByte(drive.TotalFreeSpace)));

                                        position = 3.03;
                                        service.Update(current);
                                        return;
                                    }
                                }
                            }
                            #endregion

                            #region Backup
                            //create path for file
                            position = 3.06;
                            var createFolder = CombineFileToDirectory(pathBackup, databaseName);

                            position = 3.07;
                            var folder = CreateDirectory(createFolder);

                            name = databaseName + "_" +
                                Convert2Persian(DateTime.Now, "_") + "_" +
                                DateTime.Now.Hour.ToString("00") +
                                DateTime.Now.Minute.ToString("00") +
                                DateTime.Now.Second.ToString("00");

                            position = 3.08;
                            pathBackup = CombineFileToDirectory(createFolder, name + ".bak");

                            #region Backup Database

                            position = 3.09;
                            var cnn = "Data Source=" + serverName + ";Initial Catalog=" +
                                databaseName + ";Integrated Security=true;";

                            SqlConnection sqlConnection = new SqlConnection(cnn);
                            try
                            {
                                if (CheckDirectoryExists(createFolder))
                                {
                                    position = 3.10;
                                    var cmdText = "BACKUP DATABASE [" + databaseName + "] TO  DISK = N'" +
                                        pathBackup + "' WITH NOFORMAT, NOINIT,  NAME = N'" +
                                        name + "', SKIP, REWIND, NOUNLOAD, COMPRESSION,  STATS = 10, CHECKSUM";

                                    SqlCommand SqlCom = new SqlCommand(cmdText, sqlConnection);
                                    SqlCom.CommandType = CommandType.Text;
                                    SqlCom.CommandTimeout = 700;

                                    position = 3.11;
                                    if (sqlConnection.State != ConnectionState.Open)
                                        sqlConnection.Open();

                                    position = 3.12;
                                    SqlCom.ExecuteNonQuery();
                                    //------------------- Copy to share -------------------------

                                    #region Copy the backup to the selected folder

                                    position = 3.13;
                                    desfolder = CombineFileToDirectory(pathBackup, shareFolder + "\\" + name + ".bak");

                                    try
                                    {
                                        if (domainUsername != string.Empty && domain != string.Empty && domainPassword != string.Empty)
                                        {
                                            AppDomain.CurrentDomain.SetPrincipalPolicy(System.Security.Principal.PrincipalPolicy.WindowsPrincipal);
                                            using (new NetworkAuth(domainUsername, domain, domainPassword))
                                            {
                                                position = 3.14;
                                                System.IO.File.Copy(pathBackup, desfolder, true);

                                                copytoshare = true;
                                                result += "Copy to share folder " + shareFolder + " ";
                                            }
                                        }
                                        else if (shareFolder != string.Empty)
                                        {
                                            position = 3.15;
                                            System.IO.File.Copy(pathBackup, desfolder, true);

                                            copytoshare = true;
                                            result += "Copy to local share " + shareFolder + " ";
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        result += position + "* " + ex.Message;
                                    }
                                    #endregion
                                }
                                else
                                {
                                    result += "Path " + createFolder + " was not found";
                                }

                                position = 3.16;

                                #region Backup result

                                if (CheckFileExists(pathBackup))
                                {
                                    position = 3.17;
                                    long length = new System.IO.FileInfo(pathBackup).Length;

                                    position = 3.18;
                                    result += "File Size: " + ConvertToMegaByte(length) + " MB " + Environment.NewLine;
                                    result += "Backup completed successfully.";

                                    if (!current.Contains("ymb_backuppath"))
                                        current.Attributes.Add("ymb_backuppath", pathBackup);
                                    else
                                        current["ymb_backuppath"] = pathBackup;

                                    position = 3.19;
                                    if (!current.Contains("ymb_name"))
                                    {
                                        if (copytoshare)
                                            current.Attributes.Add("ymb_name", desfolder);
                                    }
                                    else
                                    {
                                        if (copytoshare)
                                            current["ymb_name"] = desfolder;
                                    }
                                }
                                else
                                {
                                    result += "File " + name + " was not found";
                                }
                                #endregion
                            }
                            catch (Exception ex)
                            {
                                result += position + "* " + ex.Message;
                            }
                            finally
                            {
                                sqlConnection.Close();
                            }
                            #endregion

                            position = 3.21;
                            if (!current.Contains("ymb_backupresult"))
                                current.Attributes.Add("ymb_backupresult", result);
                            else
                                current["ymb_backupresult"] = result;

                            position = 3.22;
                            service.Update(current);

                            #endregion
                        }
                    }
                    else if (context.MessageName.Equals("Delete", StringComparison.InvariantCultureIgnoreCase))
                    {
                        #region Delete Plan

                        //in delete message, delete the backup file
                        position = 5.00;
                        retCurrent = (context.InputParameters.Contains("Target") && context.InputParameters["Target"] != null) ? (EntityReference)context.InputParameters["Target"] : null;

                        position = 5.01;
                        current = service.Retrieve(retCurrent.LogicalName, retCurrent.Id, new ColumnSet("ymb_backuppath", "statuscode"));

                        position = 5.02;
                        var path = current.Contains("ymb_backuppath") ? current["ymb_backuppath"].ToString() : "";

                        position = 5.04;
                        if (CheckFileExists(path))
                        {
                            position = 5.05;
                            System.IO.File.Delete(path);
                        }
                        else
                            MessageBox(OperationStatus.Canceled, "The backup file was not found.");

                        #endregion
                    }
                }
            }
            catch (InvalidPluginExecutionException iex)
            {
                throw MessageBox(iex.Status, iex.Message, this, position);
            }
            catch (Exception ex)
            {
                throw MessageBox(OperationStatus.Failed, ex.ToString(), this, position);
            }
        }

        protected long ConvertToMegaByte(long bytes)
        {
            return (bytes / 1024) / 1024;
        }

    }
}
