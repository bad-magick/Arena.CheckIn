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
    public class CustomArenaMailQueue : AgentWorker
    {
        private const int STATE_OK = 0;
        private string _publicUrl = string.Empty;
        private string _enableLogging = "false";
        private string _rootURLFolder = "arena";
        private int _maxDaysOld = 30;
        private int OrganizationId;
        public int notifyTag = -1;
        private ArenaSendMail arenaMail;
        private WorkerResultStatus globalStatus = WorkerResultStatus.Ok;
        private bool error;

        [TextSetting("Public URL", "URL to the public site (used to ensure image links in the e-mail are correct.)",
            true), Description("URL to the public site (used to ensure image links in the e-mail are correct.)")]
        public string PublicUrl
        {
            get { return this._publicUrl; }
            set { this._publicUrl = value; }
        }

        [BooleanSetting("Enable Logging", "Enables detailed logging of the SMTP transactions.", true, false),
         Description("Enables detailed logging of the SMTP transactions.")]
        public string EnableLogging
        {
            get { return this._enableLogging; }
            set { this._enableLogging = value; }
        }

        [TextSetting("Root URL Folder",
            "Folder name to look for to prefix with Public URL (used to ensure images links in the e-mail are correct)."
            , false),
         Description(
             "Folder name to look for to prefix with Public URL (used to ensure images links in the e-mail are correct)."
             )]
        public string RootURLFolder
        {
            get { return this._rootURLFolder; }
            set { this._rootURLFolder = value; }
        }

        [NumericSetting("Max Days Old", "Maximum number days old an email can be in order for the agent to send.", true)
        , Description("Maximum number days old an email can be in order for the agent to send.")]
        public int MaxDaysOld
        {
            get { return this._maxDaysOld; }
            set { this._maxDaysOld = value; }
        }

        public WorkerResultStatus ProcessQueue(out string message, out int state)
        {
            WorkerResultStatus result = WorkerResultStatus.Ok;
            message = string.Empty;
            state = 0;
            this.arenaMail = new ArenaSendMail();
            bool isEmailEnabled = SMSHelper.IsSMSViaEmailEnabled();
            this.OrganizationId = ArenaContext.Current.Organization.OrganizationID;
            Organization organization = new Organization(this.OrganizationId);
            Credentials credentials = new Credentials(organization.OrganizationID, CredentialType.SMTP);
            string mailServer = organization.Settings[OrganizationSetting.EMAIL_SERVER_KEY];
            string username = credentials.Username;
            string password = credentials.Password;
            try
            {
                PersonCommunicationCollection personCommunicationCollection = new PersonCommunicationCollection();
                personCommunicationCollection.LoadByStatusAndMaxTime("Pending", DateTime.Now.AddHours(-1.0));
                personCommunicationCollection.LoadByStatus("Queued");
                personCommunicationCollection.LoadByStatus("FAILED");
                foreach (PersonCommunication personCommunication in personCommunicationCollection)
                {
                    //PersonCommunication personCommunication = (PersonCommunication)enumerator;
                    Communication communication = new Communication(personCommunication.CommunicationID);
                    this.error = false;
                    if (communication.Status == ApprovalStatus.Created &&
                        communication.SendFutureDateTime == default(DateTime))
                    {
                        communication.Status = ApprovalStatus.Approved;
                        communication.SendFutureDateTime = DateTime.Now;
                        communication.SendWhen = CommunicationSendWhen.Now;
                    }
                    if (communication.Status == ApprovalStatus.Approved)
                    {
                        if ((communication.SendWhen == CommunicationSendWhen.Now &&
                             communication.DateCreated < DateTime.Now.AddDays((-(double) this.MaxDaysOld))) ||
                            (communication.SendWhen == CommunicationSendWhen.Future &&
                             communication.SendFutureDateTime < DateTime.Now.AddDays((-(double) this.MaxDaysOld))))
                        {
                            new PersonCommunicationData().SavePersonCommunication(personCommunication.CommunicationID,
                                                                                  personCommunication.PersonID,
                                                                                  DateTime.Now,
                                                                                  "Failed -- Older than configured Max Days Old",
                                                                                  ArenaContext.Current.Organization.
                                                                                      OrganizationID);
                        }
                        else
                        {
                            if (communication.SendWhen == CommunicationSendWhen.Now ||
                                (communication.SendWhen == CommunicationSendWhen.Future &&
                                 communication.SendFutureDateTime < DateTime.Now))
                            {
                                if (!communication.SendAttempted)
                                {
                                    communication.SendAttempted = true;
                                    communication.Save(communication.CreatedBy);
                                }
                                bool isSupportedCommunication;
                                if (communication.CommunicationMedium == CommunicationMedium.Email)
                                {
                                    isSupportedCommunication = true;
                                }
                                else
                                {
                                    if (communication.CommunicationMedium == CommunicationMedium.SMS && isEmailEnabled)
                                    {
                                        isSupportedCommunication = true;
                                    }
                                    else
                                    {
                                        if (communication.CommunicationMedium != CommunicationMedium.SMS)
                                        {
                                            throw new InvalidOperationException(
                                                "Unsupported medium for the ArenaMailQueue agent worker: " +
                                                communication.CommunicationMedium);
                                        }
                                        isSupportedCommunication = false;
                                    }
                                }
                                if (isSupportedCommunication)
                                {
                                    if (communication.CommunicationMedium == CommunicationMedium.Email &&
                                        personCommunication.Emails.Active.Count == 0)
                                    {
                                        new PersonCommunicationData().SavePersonCommunication(
                                            personCommunication.CommunicationID, personCommunication.PersonID,
                                            DateTime.Now, "Failed -- No Email Address",
                                            ArenaContext.Current.Organization.OrganizationID);
                                        continue;
                                    }
                                    if (communication.CommunicationMedium == CommunicationMedium.SMS &&
                                        personCommunication.Phones.FirstActiveSMSViaEmail == string.Empty)
                                    {
                                        new PersonCommunicationData().SavePersonCommunication(
                                            personCommunication.CommunicationID, personCommunication.PersonID,
                                            DateTime.Now, "Failed -- No SMS Email Address",
                                            ArenaContext.Current.Organization.OrganizationID);
                                        continue;
                                    }
                                    try
                                    {
                                        StringBuilder stringBuilder =
                                            new StringBuilder(this.FixUrl(communication.HtmlMessage));
                                        new StringBuilder(communication.TextMessage);
                                        EmailMessage emailMessage = new EmailMessage(mailServer);
                                        emailMessage.BeforeEmailSend += this.msg_BeforeEmailSend;
                                        emailMessage.MergedRowSent += this.msg_MergedRowSent;
                                        if (username != string.Empty)
                                        {
                                            emailMessage.Username = username;
                                            emailMessage.Password = password;
                                        }
                                        if (Convert.ToBoolean(this.EnableLogging))
                                        {
                                            emailMessage.Logging = true;
                                            emailMessage.LogPath = Directory.GetCurrentDirectory() + "//ArenaMailQueue_" +
                                                                   base.Name.Replace(" ", "") + ".log";
                                        }
                                        emailMessage.FromName = communication.SenderName;
                                        emailMessage.FromAddress = communication.SenderEmail;
                                        if (communication.ReplyTo.Trim() != string.Empty)
                                        {
                                            emailMessage.ReplyTo = communication.ReplyTo;
                                        }
                                        if (communication.CC.Trim() != string.Empty)
                                        {
                                            emailMessage.Cc = communication.CC;
                                        }
                                        if (communication.BCC.Trim() != string.Empty)
                                        {
                                            emailMessage.Bcc = communication.BCC;
                                        }
                                        emailMessage.Subject = communication.Subject;
                                        emailMessage.TextBodyPart = ArenaTextTools.HtmlToText(communication.TextMessage);
                                        emailMessage.HtmlBodyPart = ((communication.CommunicationMedium ==
                                                                      CommunicationMedium.Email)
                                                                         ? stringBuilder.ToString()
                                                                         : string.Empty);
                                        emailMessage.ContentTransferEncoding = MailEncoding.QuotedPrintable;
                                        emailMessage.BodyFormat = ((communication.CommunicationMedium ==
                                                                    CommunicationMedium.Email)
                                                                       ? MailFormat.Html
                                                                       : MailFormat.Text);
                                        emailMessage.ThrowException = false;
                                        emailMessage.ValidateAddress = false;
                                        if (
                                            !string.IsNullOrEmpty(
                                                organization.Settings[OrganizationSetting.EMAIL_PORT_KEY]))
                                        {
                                            emailMessage.Port =
                                                Convert.ToInt32(
                                                    organization.Settings[OrganizationSetting.EMAIL_PORT_KEY]);
                                        }
                                        bool flag3 = false;
                                        if (
                                            bool.TryParse(organization.Settings[OrganizationSetting.EMAIL_USE_SSL_KEY],
                                                          out flag3) && flag3)
                                        {
                                            SslSocket sslSocket = new SslSocket();
                                            if (
                                                !string.IsNullOrEmpty(
                                                    organization.Settings[OrganizationSetting.EMAIL_SSL_LOG_PATH_KEY]))
                                            {
                                                sslSocket.LogPath =
                                                    organization.Settings[OrganizationSetting.EMAIL_SSL_LOG_PATH_KEY];
                                                sslSocket.Logging = true;
                                            }
                                            emailMessage.LoadSslSocket(sslSocket);
                                        }
                                        PersonCommunicationCollection personCommunicationCollection2 =
                                            new PersonCommunicationCollection();
                                        switch (communication.CommunicationMedium)
                                        {
                                            case CommunicationMedium.Email:
                                                emailMessage.AddTo("##Email##");
                                                foreach (ArenaDataBlob current in communication.Attachments)
                                                {
                                                    if (current.ByteArray != null)
                                                    {
                                                        emailMessage.AddAttachment(new Attachment(current.ByteArray,
                                                                                                  current.
                                                                                                      OriginalFileName));
                                                    }
                                                }
                                                if (personCommunication.Emails.Count > 0)
                                                {
                                                    personCommunicationCollection2.Add(personCommunication);
                                                    DataTable dataTable = personCommunicationCollection2.DataTable(
                                                        null, true, false);
                                                    emailMessage.SendMailMerge(dataTable, 100, 30000);
                                                    if (this.error)
                                                    {
                                                        throw new SystemException("Mail merge was not successful");
                                                    }
                                                }
                                                break;
                                            case CommunicationMedium.SMS:
                                                emailMessage.AddTo("##SMSEmail##");
                                                if (personCommunication.Phones.FirstActiveSMSViaEmail != string.Empty)
                                                {
                                                    personCommunicationCollection2.Add(personCommunication);
                                                    DataTable dataTable2 = personCommunicationCollection2.DataTable(
                                                        null, false, true);
                                                    emailMessage.SendMailMerge(dataTable2, 100, 30000);
                                                    if (this.error)
                                                    {
                                                        throw new SystemException("Mail merge was not successful");
                                                    }
                                                }
                                                break;
                                            default:
                                                throw new InvalidOperationException("Unsupported medium: " +
                                                                                    communication.CommunicationMedium.
                                                                                        ToString());
                                        }
                                        continue;
                                    }
                                    catch (Exception ex)
                                    {
                                        result = WorkerResultStatus.Exception;
                                        message = string.Concat(new[]
                                                                    {
                                                                        "Error in Mail Queue Agent.\n\nCommunication ID: "
                                                                        ,
                                                                        personCommunication.CommunicationID.ToString(),
                                                                        "\nPerson ID: ",
                                                                        personCommunication.PersonID.ToString(),
                                                                        "\nName: ",
                                                                        personCommunication.LastName,
                                                                        ", ",
                                                                        personCommunication.FirstName,
                                                                        "\nE-mail Address(s): ",
                                                                        personCommunication.Emails.ToString(),
                                                                        "\n\nMessage:\nThis error most likely caused by an invalid e-mail address for the person above (even if the detailed message below does not say so please check that the address above is valid first).  A detailed message is below.\n\n"
                                                                        ,
                                                                        ex.Message,
                                                                        "\n\nStack Trace\n------------------------",
                                                                        ex.StackTrace
                                                                    });
                                        continue;
                                    }
                                }
                                try
                                {
                                    if (personCommunication.Phones.FirstActiveSMS == null)
                                    {
                                        new PersonCommunicationData().SavePersonCommunication(
                                            personCommunication.CommunicationID, personCommunication.PersonID,
                                            DateTime.Now, "Failed -- No SMS Enabled Phone",
                                            ArenaContext.Current.Organization.OrganizationID);
                                    }
                                    else
                                    {
                                        PersonCommunicationCollection personCommunicationCollection3 =
                                            new PersonCommunicationCollection();
                                        personCommunicationCollection3.Add(personCommunication);
                                        PersonCommunicationType personCommunicationType = new PersonCommunicationType();
                                        personCommunicationType.SendSMS(communication,
                                                                        personCommunicationCollection3.DataTable(),
                                                                        communication.CreatedBy);
                                    }
                                }
                                catch (Exception ex2)
                                {
                                    result = WorkerResultStatus.Exception;
                                    message = string.Concat(new[]
                                                                {
                                                                    "Error in Mail Queue Agent.\n\nCommunication ID: ",
                                                                    personCommunication.CommunicationID.ToString(),
                                                                    "\nPerson ID: ",
                                                                    personCommunication.PersonID.ToString(),
                                                                    "\nName: ",
                                                                    personCommunication.LastName,
                                                                    ", ",
                                                                    personCommunication.FirstName,
                                                                    "\nPhone Number(s): ",
                                                                    personCommunication.Phones.ToString(),
                                                                    "\n\nMessage:\nA detailed message is below.\n\n",
                                                                    ex2.Message,
                                                                    "\n\nStack Trace\n------------------------",
                                                                    ex2.StackTrace
                                                                });
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex3)
            {
                result = WorkerResultStatus.Exception;
                message = "Error in Mail Queue Agent.\n\n Message\n------------------------\n" + ex3.Message +
                          "\n\nStack Trace\n------------------------\n" + ex3.StackTrace;
            }
            this.arenaMail = null;
            return result;
        }

        private void msg_MergedRowSent(object sender, MergedRowSentEventArgs e)
        {
            EmailMessage emailMessage = sender as EmailMessage;
            if (emailMessage != null)
            {
                int communicationID = -1;
                int personID = -1;
                DataRow row = e.Row;
                try
                {
                    communicationID = int.Parse(row["CommunicationID"].ToString());
                    personID = int.Parse(row["PersonID"].ToString());
                }
                catch
                {
                }
                if (e.Success)
                {
                    new PersonCommunicationData().SavePersonCommunication(communicationID, personID, DateTime.Now,
                                                                          "Success",
                                                                          ArenaContext.Current.Organization.
                                                                              OrganizationID);
                    return;
                }
                if (this.arenaMail.IsDuplicate(emailMessage))
                {
                    new PersonCommunicationData().SavePersonCommunication(communicationID, personID, DateTime.Now,
                                                                          "Failed -- Duplicate",
                                                                          ArenaContext.Current.Organization.
                                                                              OrganizationID);
                    return;
                }
                this.error = true;
            }
        }

        private void msg_BeforeEmailSend(object sender, BeforeEmailSendEventArgs e)
        {
            EmailMessage emailMessage = sender as EmailMessage;
            if (emailMessage != null && this.arenaMail != null)
            {
                e.Send = this.arenaMail.ShouldSendEmail(emailMessage);
            }
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
                        WorkerResultStatus status = this.ProcessQueue(out detailedMessage, out state);
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

        public string FixUrl(string sourceString)
        {
            string text = this.RootURLFolder.ToLower();
            string text2 = sourceString;
            string text3 = text2.ToLower();
            for (int i = text3.IndexOf("src=\"/" + text + "/");
                 i >= 0;
                 i = text3.IndexOf("src=\"/" + text + "/", i + 6 + text.Length))
            {
                text2 = string.Concat(new[]
                                          {
                                              text2.Substring(0, i),
                                              "src=\"/",
                                              text,
                                              "/",
                                              text2.Substring(i + 7 + text.Length)
                                          });
            }
            if (text2.Contains("src=\"/" + text + "/"))
            {
                text2 = text2.Replace("src=\"/" + text + "/", string.Format("src=\"{0}/" + text + "/", this.PublicUrl));
            }
            else
            {
                text2 = text2.Replace("src=\"/", string.Format("src=\"{0}/", this.PublicUrl));
            }
            return text2;
        }
    }
}