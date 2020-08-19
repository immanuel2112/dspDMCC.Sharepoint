using CranBerry.Framework;
using CranBerry.Framework.Data;
using CranBerry.Framework.Plugins;
using System;
using System.Data;

namespace dspDMCC.Sharepoint
{
    [PluginInstaller(
        DisplayName = "DSP Sharepoint Integration - Python API",
        Description = "Plugin to integrate dspDMCC load manager webapp to upload files to secure sharepoint via Python API.",
        DataRowContract = typeof(SharepointPythonAPI.MyDataRowContract))
    ]

    public class SharepointPythonAPI : Plugin
    {
        protected override void OnExecute()
        {
            try
            {
                Log.Information(this, "Executing dspDMCC.Sharepoint process for load manager entry id: " + this.PageData.ID);
                IDataService sys = Host.GetDataService(Host.Page.DataSourceID);

                DataTable dt = sys.GetDataTable("SELECT [ServerAddress], [UserID], [Password] FROM CranSoft.dbo.DataSource WHERE [Database] = 'Cransoft'");
                string dbServer = dt.Rows[0]["ServerAddress"].ToString();
                string login = dt.Rows[0]["UserID"].ToString();
                string password = dt.Rows[0]["Password"].ToString();
                
                string cParams = dbServer + " " + login + " " + password + " " + this.PageData.ID;
                
                string exePath = sys.ExecuteScalar<string>("SELECT ExePath FROM dspDMCC.dbo.ztSharepoint");

                Log.Information(this, "Executing dspDMCC.Sharepoint process for load manager dbserver: " + dbServer);
                Log.Information(this, "Executing dspDMCC.Sharepoint process for load manager login: " + login);
                /// Log.Information(this, "Executing dspDMCC.Sharepoint process for load manager password: " + password);
                Log.Information(this, "Executing dspDMCC.Sharepoint process for load manager cParams: " + cParams);
                Log.Information(this, "Executing dspDMCC.Sharepoint process for load manager exePath: " + exePath);
                Log.Information(this, "Before calling ADMSharepointPyAPI script");
                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo.FileName = string.Concat(exePath);
                process.StartInfo.Arguments = cParams;
                process.Start();
                process.WaitForExit();
                process.Close();
                Log.Information(this, "After calling ADMSharepointPyAPI script");

            }
            catch (Exception ex)
            {
                Log.Error(this, "Error in dspDMCC.Sharepoint process: " + ex.Message);
                throw new Exception(string.Format("Error in uploading process"));
            }
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
