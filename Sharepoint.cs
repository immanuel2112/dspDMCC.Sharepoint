using CranBerry.Framework;
using CranBerry.Framework.Plugins;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace dspDMCC.Sharepoint
{
    [PluginInstaller(
        DisplayName = "DSP Sharepoint Integration",
        Description = "Plugin to integrate dspDMCC load manager webapp to upload files to secure sharepoint.",
        DataRowContract = typeof(Sharepoint.MyDataRowContract))
    ]

    public class Sharepoint : Plugin
    {
        private System.Diagnostics.EventLog eventLog = new System.Diagnostics.EventLog("Application");
        private ADMSharepointValue admSPValue = new ADMSharepointValue();

        protected override void OnExecute()
        {
            try
            {
                eventLog.Source = "dspDMCC.Sharepoint";
                eventLog.WriteEntry("Executing dspDMCC.Sharepoint process for load manager entry id: " + this.PageData.ID);

                Log.Information(this, "Executing dspDMCC.Sharepoint process for load manager entry id: " + this.PageData.ID);
                admSPValue.LoadManagerId = this.PageData.ID;
                admSPValue.Sys = Host.GetDataService(Host.Page.DataSourceID);

                // Get Sharepoint configuration values from DMCC ztSharepoint table.
                GetSharepointValues();

                if (admSPValue.SharepointValue != null)
                {
                    // Fetch load manager attributes
                    GetLoadManagerDetails();

                    // Fetch Target Report details
                    List<ADMTargetReportValue> targetReportValues = GetTargetReportDetails();
                    admSPValue.TargetReportValues = targetReportValues;

                    // Upload files to sharepoint
                    UploadReports();
                }

                Log.Information(this, "Execution Completed - dspDMCC.Sharepoint process for load manager entry id: " + this.PageData.ID);
 
            }
            catch (Exception ex)
            {
                Log.Error(this, ""+ex);
                eventLog.WriteEntry("" + ex, EventLogEntryType.Error);
                throw new Exception(string.Format("" + ex));
            }
        }

        private void GetSharepointValues()
        {
            SharepointValue spValue = null;
            try
            {
                DataTable dt = admSPValue.Sys.GetDataTable("SELECT TOP 1  [ID], [Name], [SiteURL], [SiteID], [ClientID], [ClientSecret], [FolderURL], [BaseFolderName] FROM [dspDMCC].[dbo].[ztSharepoint]");
                if (dt != null)
                {
                    spValue = new SharepointValue();
                    spValue.Id = dt.Rows[0]["ID"].ToString();
                    spValue.Name = dt.Rows[0]["Name"].ToString();
                    spValue.SiteURL = dt.Rows[0]["SiteURL"].ToString();
                    spValue.SiteID = dt.Rows[0]["SiteID"].ToString();
                    spValue.ClientID = dt.Rows[0]["ClientID"].ToString();
                    spValue.ClientSecret = dt.Rows[0]["ClientSecret"].ToString();
                    spValue.FolderURL = dt.Rows[0]["FolderURL"].ToString();
                    spValue.BaseFolderName = dt.Rows[0]["BaseFolderName"].ToString();
                }

            }
            catch (Exception ex)
            {
                Log.Error(this, "" + ex);
                eventLog.WriteEntry("" + ex, EventLogEntryType.Error);
                throw new Exception(string.Format("" + ex));
            }
            admSPValue.SharepointValue = spValue;
        }


        private void GetLoadManagerDetails()
        {
            try
            {
                DataTable dt = admSPValue.Sys.GetDataTable("SELECT ztLoadManagerTargetItemDetails.[ID], webConsoleAllTargetStructureSel.[Wave], webConsoleAllTargetStructureSel.[ProcessArea]" +
                                                                    ", webConsoleAllTargetStructureSel.[Object], webConsoleAllTargetStructureSel.[Target], ztLoadManagerTargetItemDetails.[WaveProcessAreaObjectTargetID]" +
                                                                    ", ztLoadManagerTargetItemDetails.[LoadCycle], ztLoadType.[LoadType], ztLoadManagerTargetItemDetails.[Version], ztParam.ReportPath " +
                                                            "FROM [dspDMCC].[dbo].[ztLoadManagerTargetItemDetails] AS ztLoadManagerTargetItemDetails " +
                                                                "INNER JOIN [dspDMCC].[dbo].[ztLoadType] AS ztLoadType " +
                                                                "ON ztLoadManagerTargetItemDetails.LoadType = ztLoadType.ID " +
                                                                "INNER JOIN [dspDMCC].[dbo].webConsoleAllTargetStructureSel as webConsoleAllTargetStructureSel " +
                                                                "ON webConsoleAllTargetStructureSel.WaveProcessAreaObjectTargetID = ztLoadManagerTargetItemDetails.WaveProcessAreaObjectTargetID " +
                                                                "CROSS JOIN Console.dbo.ztParam as ztParam " +
                                                            "WHERE ztLoadManagerTargetItemDetails.ID = " + admSPValue.LoadManagerId);
                if (dt != null)
                {
                    admSPValue.Wave = dt.Rows[0]["Wave"].ToString();
                    admSPValue.Processarea = dt.Rows[0]["ProcessArea"].ToString();
                    admSPValue.Object = dt.Rows[0]["Object"].ToString();
                    admSPValue.Target = dt.Rows[0]["Target"].ToString();
                    admSPValue.WaveProcessareaObjectTargetID = dt.Rows[0]["WaveProcessAreaObjectTargetID"].ToString();
                    admSPValue.LoadCycle = dt.Rows[0]["LoadCycle"].ToString();
                    admSPValue.InitialDelta = dt.Rows[0]["LoadType"].ToString();
                    admSPValue.Version = dt.Rows[0]["Version"].ToString();
                }

            }
            catch (Exception ex)
            {
                Log.Error(this, "" + ex);
                eventLog.WriteEntry("" + ex, EventLogEntryType.Error);
                throw new Exception(string.Format("" + ex));
            }
         }


        private List<ADMTargetReportValue> GetTargetReportDetails()
        {
            List<ADMTargetReportValue> targetReportValues = null;
            try
            {
                DataTable dt = admSPValue.Sys.GetDataTable("SELECT  WaveProcessAreaObjectTargetReportID, TargetReport, ReportType, FileLocation, SegmentByField " +
                                                           "FROM    dspDMCC.dbo.webConsoleTargetReportStructureSel " +
                                                           "WHERE   CAST(WaveProcessAreaObjectTargetID AS NVARCHAR(50)) = '" + admSPValue.WaveProcessareaObjectTargetID + "' " +
                                                           "        AND RecordCount > 0");
                if (dt != null)
                {
                    targetReportValues = new List<ADMTargetReportValue>();
                    foreach (DataRow row in dt.Rows)
                    {
                        ADMTargetReportValue admTargetReportValue = new ADMTargetReportValue();
                        admTargetReportValue.TargetReportID = row["WaveProcessAreaObjectTargetReportID"].ToString();
                        admTargetReportValue.TargetReport = row["TargetReport"].ToString();
                        admTargetReportValue.TargetReportType = row["ReportType"].ToString();
                        admTargetReportValue.TargetReportLocation = row["FileLocation"].ToString();
                        admTargetReportValue.TargetReportSegmentByField = row["SegmentByField"].ToString();

                        targetReportValues.Add(admTargetReportValue);
                    }
                }
                admSPValue.TargetReportValues = targetReportValues;

                // Fetch Segmented Target Report details 
                if (admSPValue.TargetReportValues != null)
                {
                    foreach (ADMTargetReportValue value in admSPValue.TargetReportValues)
                    {
                        if (value.TargetReportSegmentByField != null)
                        {
                            dt = admSPValue.Sys.GetDataTable("SELECT  SegmentByValue, FileLocation " +
                                                     "FROM    DSW.dbo.ttWaveProcessAreaObjectTargetReportSegment " +
                                                     "WHERE CAST(WaveProcessAreaObjectTargetReportID AS NVARCHAR(50)) = '" + value.TargetReportID + "'");
                            if (dt != null)
                            {
                                List<ADMTargetReportSegmentValue> targetReportSegmentValues = new List<ADMTargetReportSegmentValue>();
                                foreach (DataRow row in dt.Rows)
                                {
                                    ADMTargetReportSegmentValue admTargetReportSegmentValue = new ADMTargetReportSegmentValue();
                                    admTargetReportSegmentValue.TargetReportSegmentByFieldValue = row["SegmentByValue"].ToString();
                                    admTargetReportSegmentValue.TargetReportSegmentLocation = row["FileLocation"].ToString();

                                    targetReportSegmentValues.Add(admTargetReportSegmentValue);
                                }
                                value.TargetReportSegmentValues = targetReportSegmentValues;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(this, "" + ex);
                eventLog.WriteEntry("" + ex, EventLogEntryType.Error);
                throw new Exception(string.Format("" + ex));
            }
            return targetReportValues;
        }


        private void UploadReports()
        {
            try
            {
                string siteURL = admSPValue.SharepointValue.SiteURL + "sites/" + admSPValue.SharepointValue.SiteID;

                using (var ctx = new AuthenticationManager().GetAppOnlyAuthenticatedContext(siteURL, admSPValue.SharepointValue.ClientID, admSPValue.SharepointValue.ClientSecret))
                {
                    ctx.Load(ctx.Web);
                    ctx.ExecuteQuery();

                    // Create folder structure
                    string newPath = CreateFolder(ctx);
                    
                    // Upload reports without segment by configuration
                    List<ADMTargetReportValue> targetReportValues = admSPValue.TargetReportValues;

                    if (targetReportValues != null)
                    {
                        foreach (ADMTargetReportValue reportValue in targetReportValues)
                        {
                            string filePath = reportValue.TargetReportLocation;
                            string fileName = reportValue.TargetReport + ".xlsx";
                            string segmentByField = reportValue.TargetReportSegmentByField;
                            List<ADMTargetReportSegmentValue> targetSegmentReportValues = reportValue.TargetReportSegmentValues;

                            if (filePath != null && filePath.Length > 0)
                            {
                                UploadFile(filePath, fileName, newPath, ctx, null);
                            }

                            if (segmentByField != null && targetSegmentReportValues != null)
                            {
                                foreach (ADMTargetReportSegmentValue segmentValue in targetSegmentReportValues)
                                {
                                    filePath = segmentValue.TargetReportSegmentLocation;
                                    string segmentByValue = segmentValue.TargetReportSegmentByFieldValue;
                                    
                                    if (filePath != null)
                                    {
                                        UploadFile(filePath, fileName, newPath, ctx, segmentByValue);
                                    }
                                }
                            }
                        }
                    }
                };
            }
            catch (Exception ex)
            {
                Log.Error(this, "" + ex);
                eventLog.WriteEntry("" + ex, EventLogEntryType.Error);
                throw new Exception(string.Format("" + ex));
            }
        }


        private void UploadFile(string srcFilePath, string fileName, string spPath, ClientContext ctx, string segmentByValue)
        {
            try
            {
                Web web = ctx.Web;
                if (segmentByValue != null)
                {
                    fileName = fileName.Replace(".xlsx", "_" + segmentByValue + ".xlsx");
                }
                ResourcePath folderPath = ResourcePath.FromDecodedUrl(spPath + "/" + fileName);
                Folder parentFolder = web.GetFolderByServerRelativePath(folderPath);

                byte[] fileData = null;

                using (FileStream fs = System.IO.File.OpenRead(srcFilePath))
                {
                    using (BinaryReader binaryReader = new BinaryReader(fs))
                    {
                        fileData = binaryReader.ReadBytes((int)fs.Length);
                    }
                }

                FileCollectionAddParameters fileAddParameters = new FileCollectionAddParameters();
                fileAddParameters.Overwrite = true;
                using (MemoryStream contentStream = new MemoryStream(fileData))
                {
                    // Add a file
                    Microsoft.SharePoint.Client.File addedFile = parentFolder.Files.AddUsingPath(folderPath, fileAddParameters, contentStream);

                    // Select properties of added file to inspect
                    ctx.Load(addedFile, f => f.UniqueId, f1 => f1.ServerRelativePath);

                    // Perform the actual operation
                    ctx.ExecuteQuery();

                    // Print the results
                    // eventLog.WriteEntry("Added File [ServerRelativePath:" + addedFile.ServerRelativePath.DecodedUrl + "]");
                }
            }
            catch (Exception ex)
            {
                Log.Error(this, "" + ex);
                eventLog.WriteEntry("" + ex, EventLogEntryType.Error);
                throw new Exception(string.Format("" + ex));
            }
        }


        private string CreateFolder(ClientContext ctx)
        {
            string newPath;
            try
            {

                string folderURL = admSPValue.SharepointValue.FolderURL;
                string folderName = admSPValue.SharepointValue.BaseFolderName;

                // Check for existence for Base folder and create if it doesn't exists
                newPath = CreateFolderUtility(ctx, folderURL, folderName);

                // Check for existence for Wave folder and create if it doesn't exists
                newPath = CreateFolderUtility(ctx, newPath, admSPValue.Wave);

                // Check for existence for Load Cycle folder and create if it doesn't exists
                newPath = CreateFolderUtility(ctx, newPath, admSPValue.LoadCycle);

                // Check for existence for Process Area folder and create if it doesn't exists
                newPath = CreateFolderUtility(ctx, newPath, admSPValue.Processarea);

                // Check for existence for Object folder and create if it doesn't exists
                newPath = CreateFolderUtility(ctx, newPath, admSPValue.Object);

                // Check for existence for Target folder and create if it doesn't exists
                newPath = CreateFolderUtility(ctx, newPath, admSPValue.Target);

                // Check for existence for Version folder and create if it doesn't exists
                newPath = CreateFolderUtility(ctx, newPath, 'v' + admSPValue.Version.ToString());
            }
            catch (Exception ex)
            {
                Log.Error(this, "" + ex);
                eventLog.WriteEntry("" + ex, EventLogEntryType.Error);
                throw new Exception(string.Format("" + ex));
            }
            return newPath;
        }


        private string CreateFolderUtility(ClientContext ctx, string folderURL, string folderName)
        {
            string newPath;
            try
            {
                Web web = ctx.Web;
                string folderNameNew = folderName.Replace("/", "");
                // Get the parent folder
                ResourcePath folderPath = ResourcePath.FromDecodedUrl(folderURL);
                Folder parentFolder = web.GetFolderByServerRelativePath(folderPath);

                bool folderExists = parentFolder.FolderExists(folderNameNew);
                if (!folderExists)
                {
                    // Create the parameters used to add a folder
                    ResourcePath subFolderPath = ResourcePath.FromDecodedUrl(folderURL + "/" + folderNameNew);
                    FolderCollectionAddParameters folderAddParameters = new FolderCollectionAddParameters();

                    // Add a sub folder
                    Folder addedFolder = parentFolder.Folders.AddUsingPath(subFolderPath, folderAddParameters);

                    // Select properties of added file to inspect
                    ctx.Load(addedFolder, f => f.UniqueId, f1 => f1.ServerRelativePath);

                    // Perform the actual operation
                    ctx.ExecuteQuery();

                    // Print the results
                }
                newPath = folderURL + "/" + folderNameNew;
            }
            catch (Exception ex)
            {
                Log.Error(this, "" + ex);
                eventLog.WriteEntry("" + ex, EventLogEntryType.Error);
                throw new Exception(string.Format("" + ex));
            }
            return newPath;
        }


        #region DataRowContract
        /// <summary>
        /// DataRowContract class to get data off from Load Manager item details page.
        /// </summary>
        private sealed class MyDataRowContract : DataRowContract
        {
            public String ID { get; set; }
        }

        /// <summary>
        /// Get the data row contract object.
        /// </summary> 
        private MyDataRowContract PageData
        {
            get { return (MyDataRowContract)base.DataRowContract; }
        }
    }
    #endregion
}
