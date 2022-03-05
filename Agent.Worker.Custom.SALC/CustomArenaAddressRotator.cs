#region Using

using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Text;
using AdvancedIntellect.Ssl;
using Arena.Core;
using Arena.Core.Communications;
using Arena.Core.SMS;
using Arena.DataLayer.Core;
using Arena.Enums;
using Arena.Organization;
using Arena.Portal;
using Arena.Security;
using Arena.Utility;
using aspNetEmail;

#endregion

namespace Agent.Worker.Custom.SALC
{
    [Description("Sends all pending emails from the communication object.")]
    [Serializable]
    public class CustomArenaAddressRotator : AgentWorker
    {
        private string _enableLogging = "false";
        private int _organizationId;

        [BooleanSetting("Enable Logging", "Enables detailed logging of the SMTP transactions.", true, false),
         Description("Enables detailed logging of the SMTP transactions.")]
        public string EnableLogging
        {
            get
            {
                return this._enableLogging;
            }
            set { this._enableLogging = value; }
        }

        public WorkerResultStatus Process(out string message, out int state)
        {
            WorkerResultStatus result = WorkerResultStatus.Ok;
            message = string.Empty;
            state = 0;
            bool isEmailEnabled = SMSHelper.IsSMSViaEmailEnabled();
            this._organizationId = ArenaContext.Current.Organization.OrganizationID;
            Organization organization = new Organization(this._organizationId);

            string addressFields = organization.Settings["Rotation Fields"];
            string[] fieldIDs = addressFields.Split(',');

            
           
            try
            {
                for (int i = 0 - 1; i < (int)(fieldIDs.Length / 3); i++)
                {
                    int type = int.Parse(fieldIDs[3 * i + 0]);
                    int date = int.Parse(fieldIDs[3 * i + 1]);
                    int year = int.Parse(fieldIDs[3 * i + 2]);

                    //PersonAttribute personAttribute = 
                }
            }
            catch (Exception ex3)
            {
                result = WorkerResultStatus.Exception;
                message = "Error in Mail Queue Agent.\n\n Message\n------------------------\n" + ex3.Message +
                          "\n\nStack Trace\n------------------------\n" + ex3.StackTrace;
            }
            return result;
        }

        public override WorkerResult Run(bool previousWorkersActive)
        {
            WorkerResult result;
            try
            {
                if (Convert.ToBoolean(this.Enabled))
                {
                    if (this.RunIfPreviousWorkersActive || !previousWorkersActive)
                    {
                        string detailedMessage;
                        int state;
                        WorkerResultStatus status = this.Process(out detailedMessage, out state);
                        result = new WorkerResult(state, status, string.Format(base.Description, new object[0]),
                                                  detailedMessage);
                    }
                    else
                    {
                        result = new WorkerResult(0, WorkerResultStatus.Abort,
                                                  string.Format(base.Description, new object[0]),
                                                  "Did not run because previous worker instance still active.");
                    }
                }
                else
                {
                    result = new WorkerResult(0, WorkerResultStatus.Abort,
                                              string.Format(base.Description, new object[0]),
                                              "Did not run because worker not enabled.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return result;
        }

    }
}