using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Configuration;
using System.Collections;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using OpenPop.Mime;
using OpenPop.Mime.Header;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using OpenPop.Common.Logging;
using Message = OpenPop.Mime.Message;
using System.Reflection;
using Takata.Global.ServerComponents.BusinessServices.Components;
using System.Globalization;
using Microsoft.Exchange.WebServices.Data;
using System.Net;
using System.Text.RegularExpressions;

namespace TNRoutingEmailProcessor
{
    class Program
    {
        #region Constants

        private const string strMessage1 = "DO NOT CHANGE THE EMAIL SUBJECT.";
        private const string strMessage2 = "Please type any comments here...";
        private const string strMessage3 = "P Think Green - Please consider the environment before printing this email.";
        private const string strMessage4 = "The information in this email and attachments hereto may contain legally privileged, proprietary or confidential information that is intended for a particular recipient. If you are not the intended recipient(s), or the employee or agent responsible for delivery of this message to the intended recipient(s), you are hereby notified that any disclosure, copying, distribution, retention or use of the contents of this e-mail information is prohibited and may be unlawful. When addressed to Joyson Safety Systems customers or vendors, any information contained in this e-mail is subject to the terms and conditions in the governing contract, if applicable. If you have received this communication in error, please immediately notify us by return e-mail, permanently delete any electronic copies of this communication and destroy any paper copies.";
        private const string strMessage5 = "________________________________";
        private const string strChar1 = "\r";
        private const string strChar2 = "\n";
        private static string strEnvironment = String.Empty;
        private const string strApplicationName = "TNRoutingEmailProcessor";
        private const int iLogLevel = 1; //1 is the lowest Logging
        private static int iCurLogLevelSet = 1;
        private static int iRunCtr = 1;
        private static int iMaxTries = 100;
        private static Boolean blnConnectError = false;
        private const string APPR_GUID = "1CC4D208-3998-4D2B-9EF9-1CE684E1A562";
        private const string REJ_GUID = "03226954-2FB1-40EB-8F99-FDADE5D893A7";
        private const string CLAR_GUID = "B1F1C49C-AF7B-4683-B54C-D045CEC93C56";

        private const string RT_PKGREJECTED = "Rejected";
        private const string RT_PKGAPPROVED = "Approved";
        private const string RT_PKGINPROCESS = "In-Process";
        private const string RT_PKGCREATED = "Draft";
        private const string RT_PKGSEEKCLARIFICATION = "Seek Clarification";
        private const string RT_PKGAPPROVEDBYASSIGNEE = "Awaiting final Approval";
        private const string RT_ASSIGNEESHUFFLINGUP = "ShufflingUp";
        private const string RT_ASSIGNEESHUFFLINGDOWN = "ShufflingDown";
        private const string RT_ASSIGNEEREINITROUTE = "ReInitRoute";
        private const string RT_ASSIGNEECANCEL = "CancelSave";
        private const string RT_ASSIGNEEROTHER = "Other";
        private const string RT_INITIATORROUTEORDER = "0";
        private const string RT_INITIATORAUTH = "InitiatorAuth";
        private const string RT_USERHASGIVENDISPOSITION = "1";
        private const string RT_USERISNOTNEXTASSIGNEE = "2";
        private const string RT_USERISALTERNATE = "3";
        private const string RT_USERISNOTASSIGNEE = "4";
        private const string RT_USERAUTH = "5";
        private const string RT_USERISNOTCURRASSIGNEE = "6";
        private const string RT_USERNEEDTOWAIT = "7";
        private const string RT_USERISINITIATOR = "8";
        private const string RT_STYLESHEET = "/SharedResources/Styles/GlobalStyle.css";
        private const string RT_PKG_SECURED = "EF21BA51-1D58-41C5-8D98-B34DE7ABBD86";
        private const string RT_PKG_PUBLIC = "D5206072-6C52-48F5-B4A6-6F1981E1B6D3";
        private const string RT_PKG_REJECTED_ID = "C4335908-B0E6-4920-90FF-03358391B2FF";

        private static bool ProcessComplete = false;
        private static string strHDEMailID = "TasTechSupport@Takata.com";
        private static string strTakataNetEmail = "TakataNet@Takata.com";
        private static string strTasTechSupportEmail = "TasTechSupport@Takata.com";

        private const string strStarted = "Started";
        private const string strEnded = "Ended";
        private const string strError = "Error";
        private const string strIN = " in ";
        private const string strParameters = "Parameters:";

        //Takata Email Approval
        private const string STR_ROUTINGDESC = "Package Description:";
        private const string STR_ROUTINGTYPE = "Package Type:";
        private const string STR_ROUTING = "Routing: ";

        private const string STR_DISPOSITION = " Status ";
        private const string STR_ASSIGNEE = "Assignee";
        private const string STR_DISPOSITIONDATE = "Date of Disposition";
        private const string STR_COMMENTS = " Comments ";
        private const string STR_UTCDATE = " GMT";
        //private const string STR_SEPARATOR = "|";

        #endregion

        static void Main(string[] args)
        {

            strHDEMailID = ConfigurationManager.AppSettings["HDEmailID"].ToString();
            strTakataNetEmail = ConfigurationManager.AppSettings["TakataNetEmail"].ToString();
            strTasTechSupportEmail = ConfigurationManager.AppSettings["TasTechSupportEmail"].ToString(); 
            
            blnConnectError = false;
            iCurLogLevelSet = Convert.ToInt32(ConfigurationManager.AppSettings["LogLevel"]);

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            ProcessEmails_EWS();

            if ((blnConnectError == true) && (iRunCtr < 10))
            {
                blnConnectError = false;
                iRunCtr++;
                ProcessEmails_EWS();
            }

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));
        }

        /// <summary>
        /// Gets all email from the inbox for EWSTest@Takata.com, inserts into the database and 
        /// processes all records from the database to set them as approved or rejected
        /// </summary>
        private static void ProcessEmails_EWS()
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            try
            {
                //bool retValInstoDB = true;
                
                while (ProcessComplete == false)
                {
                    Console.WriteLine("Email extraction started!");
                    EWSGetMessagefromExchange();
                    Console.WriteLine("Email extraction ended!");
                }
                
                Console.WriteLine("ProcessApprovalData Started!");
                ProcessApprovalData();
                Console.WriteLine("ProcessApprovalData Ended!");

                if (iCurLogLevelSet > iLogLevel)
                    AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));
            }
            catch (Exception ex)
            {
                AddLog(String.Concat(strError, strIN, MethodInfo.GetCurrentMethod().Name));
                AddLog(ex.Message);
            }
        }


        private static bool certificateValidator(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslpolicyerrors)
        {
            // We should check if there are some SSLPolicyErrors, but here we simply say that
            // the certificate is okay - we trust it.
            return true;
        }

        private static string GetUserNamefromEmail(string strEmailId)
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            CDataSet dstParams = new CDataSet();
            dstParams.Fields.Add("strEmail", strEmailId);
            Routing_bc bclRouting = new Routing_bc();
            dstParams.LoadXML(bclRouting.GetEmpDatafromEmail(dstParams.XML), false);

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

            if (dstParams.DataSets != null && dstParams.DataSets.Count > 0)
                return dstParams.DataSets[0].Fields["Assignee"].FieldValue;
            else
                return String.Empty;

        }

        private static bool UpdatePackageApprovalStatus(CDataSet dstParams, string strAction)
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            try
            {
                Routing_bc bclRouting = new Routing_bc();

                if (strAction.ToUpper() == APPR_GUID)
                    dstParams.LoadXML(bclRouting.ApprovePackage(dstParams.XML), false);
                else if (strAction.ToUpper() == REJ_GUID)
                    dstParams.LoadXML(bclRouting.RejectPackage(dstParams.XML), false);
                else if (strAction.ToUpper() == CLAR_GUID)
                    dstParams.LoadXML(bclRouting.ClarifyPackage(dstParams.XML), false);

                if (dstParams.IsErrorDataSet)
                    return false;
                else
                    return true;
            }
            catch (Exception ex)
            {
                AddLog(String.Concat(strError, strIN, MethodInfo.GetCurrentMethod().Name));
                AddLog(strParameters + dstParams.XML);
                AddLog(ex.Message);
                return false;

            }
            finally
            {
                if (iCurLogLevelSet > iLogLevel)
                    AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

            }
        }

        /// <summary>
        /// Gets all emails from the outlook365 exchange for the given UPN/email account and calls EWSInsertEmailtoDB() for each item
        /// </summary>
        /// <returns></returns>
        private static void EWSGetMessagefromExchange()
        {
            String HostName = String.Empty;
            String UserName = String.Empty;
            String Password = String.Empty;
            
            List<EmailMessage> emails = new List<EmailMessage>();
            FindItemsResults<Item> findResults;

            HostName = ConfigurationManager.AppSettings["EWSServiceUri"].ToString();
            UserName = ConfigurationManager.AppSettings["EWSUserName"].ToString();
            Password = ConfigurationManager.AppSettings["EWSPassword"].ToString();

            strEnvironment = ConfigurationManager.AppSettings["Environment"].ToString();

            int offset = 0;
            int pageSize = 50;
            bool more = true;
            bool insertSuccess = false;

            // Create a view with a page size of 50.
            ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);
            view.PropertySet = PropertySet.FirstClassProperties;
            
            ExchangeService service = new ExchangeService();

            try
            {
                ProcessComplete = false;

                //Authentication - Set specific credentials.
                service.Credentials = new NetworkCredential(UserName, Password);
                //Endpoint management Set the URL manually
                service.Url = new Uri(HostName);

                while (more)
                {
                    // Send the request to the Inbox and get the results.
#if DEBUG
                    findResults = service.FindItems(WellKnownFolderName.Drafts, view);
#else
                    findResults = service.FindItems(WellKnownFolderName.Inbox, view);
#endif
                    if (findResults.Items == null || findResults.Items.Count == 0)
                    {
                        ProcessComplete = true;
                        more = findResults.MoreAvailable;
                        Console.WriteLine("There are no emails to process");
                    }
                    else
                    {
                        Console.WriteLine("There are " + findResults.Items.Count + " emails to process");
                        //Load Properties for all emails to access email Body and other properties
                        service.LoadPropertiesForItems(findResults, PropertySet.FirstClassProperties);

                        foreach (var item in findResults.Items)
                        {
                            //emails.Add((EmailMessage)item);
                            insertSuccess = EWSInsertEmailtoDB((EmailMessage)item);
                            //RM Insert into DB, before delete
                            if (insertSuccess)
                                item.Delete(DeleteMode.MoveToDeletedItems);
                        }
                        more = findResults.MoreAvailable;
                        if (more)
                        {
                            view.Offset += pageSize;
                        }
                        else
                            ProcessComplete = true;
                    }
                }

            }
            catch (Exception ex)
            {
                ProcessComplete = false;
                AddLog(String.Concat(strError, strIN, MethodInfo.GetCurrentMethod().Name));
                AddLog(ex.Message);
                blnConnectError = true;
            }

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

        }

        private static void AddLog(string strLogMsg)
        {
            try
            {
                CDataSet dstParams = new CDataSet();
                dstParams.Fields.Add("LogMsg", strLogMsg);
                dstParams.Fields.Add("AppName", strApplicationName);

                using (Routing_bc bclRouting = new Routing_bc())
                {
                    bclRouting.AddLog(dstParams.XML);
                }
            }
            catch (Exception ex)
            {
                SendMailToHD(ex.StackTrace, "Error in AddLog", strHDEMailID, DateTime.Now.ToString(), ex.Message);
            }

        }

        
        private static string CheckDisposition(String strPackageId, string strAssignee)
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));


            CDataSet dstParams = new CDataSet();
            try
            {

                dstParams.Fields.Add("PackageId", strPackageId);
                dstParams.Fields.Add("Assignee", strAssignee);
                using (Routing_bc bclRouting = new Routing_bc())
                {
                    dstParams.LoadXML(bclRouting.CheckDisposition(dstParams.XML), false);
                }
                return dstParams.Fields["RETURN_VALUE"].FieldValue;

            }
            catch (Exception errObject)
            {
                AddLog(String.Concat(strError, strIN, MethodInfo.GetCurrentMethod().Name));
                AddLog(errObject.Message);

            }
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

            return String.Empty;
        }

        /// <summary>
        /// If error is valid, send email with appropriate error message and update email process table
        /// </summary>
        /// <param name="strEmailAddress"></param>
        /// <param name="bclRouting"></param>
        /// <param name="strPackageID"></param>
        /// <param name="m_strPkgDesc"></param>
        /// <param name="m_strPackageType"></param>
        /// <param name="errorReason"></param>
        private static void HandleErrorReason(string strEmailAddress, Routing_bc bclRouting, string strPackageID, string m_strPkgDesc, string m_strPackageType, string errorReason)
        {
            string strErrorMessage = String.Empty;
            switch (errorReason)
            {
                case RT_PKGAPPROVED: //9
                    strErrorMessage = "The package has already been Approved.";
                    AddLog(strErrorMessage);
                    break;
                case RT_PKGREJECTED: 
                    strErrorMessage = "The package has already been Rejected by another assignee/initiator";
                    AddLog(strErrorMessage);
                    break;
                case RT_USERISNOTASSIGNEE: //4
                    strErrorMessage = "You are not an assignee for this package, you could have been removed while the package was still in routing.";
                    AddLog(strErrorMessage.Replace("You", GetUserNamefromEmail(strEmailAddress)));
                    break;
                case RT_USERHASGIVENDISPOSITION: //1
                    //RM - this package is already approved by this user (from sproc logic)
                    //user has given disposition "1"
                    //dstParams = new CDataSet();
                    //dstParams.Fields.Add("PackageId", strPackageId);
                    //dstParams.Fields.Add("UserName", strSenderId);
                    //using (Routing_bc bclRouting = new Routing_bc())
                    //{
                    //    dstParams.LoadXML(bclRouting.GetActualDispositionUser(dstParams.XML), false);
                    //    //get actual disposition of current use
                    //}

                    ////check dataset
                    //if (!dstParams.IsErrorDataSet)
                    //{
                    //    string strActualDispositionUser = dstParams.DataSets[0].Fields["Actual"].FieldValue;
                    //    //set return value in strActualDispositionUser variable
                    //    AddLog("Error in Disposition #1. Not proceeding");
                    //    //user has already approved the package "1"
                    //    // this could have been approved by another user (then check the actual user and include thsi info in the email)
                    //    AddLog(strActualDispositionUser + "  has already approved the package");
                    //    return blnDonotProceed;
                    //}
                    //else
                    //{
                    //    AddLog(dstParams.DataSets["Error"].Fields["ErrDesc"].FieldValue);
                    //    return blnDonotProceed;
                    //}


                    strErrorMessage = "You have already approved this package.";
                    AddLog(strErrorMessage.Replace("You", GetUserNamefromEmail(strEmailAddress)));
                    break;
                case RT_USERISINITIATOR: //8
                    strErrorMessage = "You are the initiator.";
                    AddLog(strErrorMessage.Replace("You", GetUserNamefromEmail(strEmailAddress)));
                    break;
                case RT_USERISNOTNEXTASSIGNEE: //2
                    strErrorMessage = "There are others before you that need to approve the package.";
                    AddLog(strErrorMessage.Replace("you", GetUserNamefromEmail(strEmailAddress)));
                    break;
                case RT_USERNEEDTOWAIT: //7
                    strErrorMessage = "The previous approver has not approved the package yet.";
                    AddLog(strErrorMessage);
                    break;
                default:
                    break;
            }

            SendMailToApprover(strEmailAddress, bclRouting, strPackageID, m_strPkgDesc, m_strPackageType, strErrorMessage);

        }

        /// <summary>
        /// Identity the error with the user, to check if the program could proceed or not
        /// </summary>
        /// <param name="strAssigneeStatus"></param>
        /// <param name="strPackageId"></param>
        /// <param name="strSenderId"></param>
        /// <returns></returns>
        private static bool CheckAssigneeStatus(string strAssigneeStatus, string strPackageId, string strSenderId)
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            Boolean blnDonotProceed = false;
            Boolean blnproceed = true;

            try
            {
                switch (strAssigneeStatus)
                {
                    //check return value of checkdisposittion 
                    case RT_USERHASGIVENDISPOSITION:
                    case RT_USERISNOTNEXTASSIGNEE:
                    case RT_USERISALTERNATE:
                    case RT_USERISNOTASSIGNEE:
                    case RT_USERNEEDTOWAIT:
                    case RT_USERISINITIATOR: //User is Initiator -8
                        AddLog(strSenderId + "  cannot proceed with this package. Disposition: " + strAssigneeStatus.ToString() );
                        return blnDonotProceed;
                    //RM - not sure when this condition happens
                    //case RT_INITIATORROUTEORDER:
                    //    //- with document rejected already by another assignee/initiator
                    //    //Only Rejected Status would show show error.
                    //    CDataSet dsttemp = GetPackageDetails(strPackageId, strSenderId);
                    //    String m_strPkgStatus = String.Empty;
                    //    if ((dsttemp != null) && (dsttemp.DataSets.Count > 0))
                    //        m_strPkgStatus = dsttemp.DataSets[0].Fields["pkgStatus"].FieldValue;

                    //    if (m_strPkgStatus == "Rejected")
                    //    {
                    //        AddLog("Document rejected already by another assignee/initiator");
                    //        return blnDonotProceed;

                    //    }
                    //    break;
                    default:  //other user is valid and visible all page link
                        break;

                }

                if (iCurLogLevelSet > iLogLevel)
                    AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

                return blnproceed;
            }
            catch (Exception ErrObject)
            {
                AddLog(String.Concat(strError, strIN, MethodInfo.GetCurrentMethod().Name));
                AddLog(ErrObject.Message);
                return blnDonotProceed;
            }

        }
        private static CDataSet GetPackageDetails(string strPackageId, string strSender)
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            CDataSet dstParams = null;
            CDataSet dstReturn = null;
            try
            {
                dstParams = new CDataSet();
                dstReturn = new CDataSet();
                {
                    dstParams.Fields.Add("PackageId", strPackageId);
                    dstParams.Fields.Add("UserName", strSender);
                    using (Routing_bc bclRouting = new Routing_bc())
                    {
                        dstReturn.LoadXML(bclRouting.RetrievePackage(dstParams.XML), false);
                        //Calls GetPackageId function by passing DocId to get document packageId
                    }
                }
            }
            catch (Exception errObject)
            {
                AddLog(MethodInfo.GetCurrentMethod().Name + " " + strError);
                AddLog(errObject.Message);
            }

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

            return dstReturn;
        }

        /// <summary>
        /// Inserts an EmailMessage into the dbo.Routing_EmailMessage table in TakataNet database
        /// </summary>
        /// <param name="emailMessage">Microsoft.Exchange.WebServices.Data.EmailMessage</param>
        /// <returns>true or false</returns>
        private static bool EWSInsertEmailtoDB(EmailMessage emailMessage)
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            bool IsSuccess = false;

            string strMessage = String.Empty;
            string strSubject = String.Empty;
            string strSender = String.Empty;
            string strAction = String.Empty;
            string[] strResArr;
            string strPackageID = String.Empty;
            string strDate = DateTime.Now.ToString();
            string strSenderOrig = String.Empty;
            try
            {
                string strEnvEmail = String.Empty;

                CDataSet dsParams = new CDataSet();
                CDataSet dstUserDetails = new CDataSet();

                strMessage = emailMessage.Body;
                strSubject = emailMessage.Subject;
                strSender = emailMessage.From.Address.ToString();
                strDate = emailMessage.DateTimeCreated.ToString(); //DateCreated or DateReceived or DateSent

                strMessage = FormatComment(strMessage);
                //strMessage = FormatCommentWithTags(strMessage);

                if (strMessage.StartsWith("<html"))
                    strMessage = RemoveTags(strMessage);

                strResArr = strSubject.Split('|');
                if (strResArr.Length > 1)
                {
                    strPackageID = strResArr[0].Replace("PackageID:", String.Empty);
                    strPackageID = strPackageID.Replace(":", String.Empty);
                    strPackageID = strPackageID.Replace("FW", String.Empty);
                    strPackageID = strPackageID.Replace("RE", String.Empty);
                    strAction = strResArr[1].Replace("Action:", String.Empty);
                    strEnvEmail = strResArr[2].Replace("Env:", String.Empty);
                }
               
                strSender = GetUserNamefromEmail(strSender);

                dsParams.Clear();
                dsParams.Fields.Add("MessageBody", strMessage);
                dsParams.Fields.Add("MessageSubject", strSubject);
                dsParams.Fields.Add("Assignee", strSender);
                dsParams.Fields.Add("EmailPackageId", strPackageID.Trim());
                dsParams.Fields.Add("Comment", strMessage);
                dsParams.Fields.Add("Routing_Action", strAction);
                dsParams.Fields.Add("IsProcessed", "False");
                dsParams.Fields.Add("NumTries", "0");
                dsParams.Fields.Add("EmailDate", strDate);

                if (strEnvEmail != strEnvironment)
                {
                    AddLog("Email for a diff environment " + dsParams.XML);
                    return false;
                }
                if (strSender.Equals(String.Empty))
                {
                    AddLog("Email for sender not found " + dsParams.XML);
                    return false;
                }

                using (Routing_bc bclRouting = new Routing_bc())
                {
                    bclRouting.AddEmailMessage(dsParams.XML);
                    Console.WriteLine(String.Format("{0} is successfully inserted", strPackageID));
                    IsSuccess = true;
                }

            }
            catch (Exception ex)
            {
                IsSuccess = false;
                AddLog(String.Concat(strError, strIN, MethodInfo.GetCurrentMethod().Name));
                SendMailToHD(strMessage, strSubject, strSender, strDate, "Error in Insert Email to DB. Exception Message" + ex.Message);
            }

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));


            return IsSuccess;
        }

        /// <summary>
        /// Gets all unprocessed records from the dbo.Routing_EmailMessage table in TakataNet database
        /// </summary>
        /// <returns></returns>
        private static CDataSet GetApprovalDatafromEmail()
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            CDataSet dstParams = new CDataSet();

            try
            {
                dstParams.Fields.Add("IsProcessed", "False");
                Routing_bc bclRouting = new Routing_bc();
                dstParams.LoadXML(bclRouting.GetApprovalDatafromEmail(dstParams.XML), false);

            }
            catch (Exception ex)
            {
                AddLog(MethodInfo.GetCurrentMethod().Name + " " + strError);
                AddLog(ex.Message);
            }
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

            return dstParams;
        }

        private static void ProcessApprovalData()
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            try
            {

                string strSender;
                string strAction = String.Empty;
                string strPackageID = String.Empty;
                string strMessage = String.Empty;
                string strEmailAddress = String.Empty;
                string strDate = DateTime.Now.ToString();
                bool IsPkgUpd = false;
                bool blnProceed = false;
                string strSenderOrig = string.Empty;
                CDataSet dstAppData = null;

                //RM - set iMaxTries from app.config
                iMaxTries = Convert.ToInt32(ConfigurationManager.AppSettings["MaxTries"].ToString());

                //RM - Get all unprocessed records from Routing_EmailMessage table
                dstAppData = GetApprovalDatafromEmail();

                for (int i = 0; i < dstAppData.DataSets.Count; i++)
                {
                    CDataSet dsParams = new CDataSet();
                    CDataSet dstUserDetails = new CDataSet();
                    IsPkgUpd = false;
                    strSender = String.Empty;
                    strPackageID = String.Empty;
                    int iNumTries = 0;

                    if (dstAppData.DataSets[i].Fields["NumTries"].FieldValue != null)
                        iNumTries = Convert.ToInt32(dstAppData.DataSets[i].Fields["NumTries"].FieldValue);

                    strSender = dstAppData.DataSets[i].Fields["Assignee"].FieldValue;
                    strPackageID = dstAppData.DataSets[i].Fields["EmailPackageId"].FieldValue;
                    strMessage = dstAppData.DataSets[i].Fields["Comment"].FieldValue;
                    strAction = dstAppData.DataSets[i].Fields["Routing_Action"].FieldValue;
                    strEmailAddress = dstAppData.DataSets[i].Fields["EmailAddress"].FieldValue;

                    dstUserDetails.Fields.Add("Assignee", strSender);
                    dstUserDetails.Fields.Add("PackageId", strPackageID.Trim());
                    dstUserDetails.Fields.Add("Comment", strMessage);

                    if (iCurLogLevelSet > iLogLevel)
                        AddLog("Parameters(dstUserDetails):" + dstUserDetails.XML);

                    using (Routing_bc bclRouting = new Routing_bc()) {

                        //RM - Get package details for each email
                        CDataSet dstReturn = GetPackageDetails(strPackageID, strSender);
                        if (dstReturn.DataSets[0].Fields["SoftError"] == null || dstReturn.DataSets[0].Fields["SoftError"].FieldValue != "True")
                        {
                            string m_strPkgDesc = dstReturn.DataSets[0].Fields["PackageDesc"].FieldValue;
                            string m_strPackageTypeID = dstReturn.DataSets[0].Fields["PackageTypeID"].FieldValue;
                            string m_strPackageType = dstReturn.DataSets[0].Fields["PackageTypeValue"].FieldValue;
                            string m_status = dstReturn.DataSets[0].Fields["pkgStatus"].FieldValue;

                            iNumTries = iNumTries + 1; //Incrementing iNumTries for the current processing

                            //if already approved or rejected send email to approver, set processed to true in DB
                            if (String.Compare(m_status, RT_PKGREJECTED, true) == 0 || String.Compare(m_status, RT_PKGAPPROVED, true) == 0) 
                            {
                                HandleErrorReason(strEmailAddress, bclRouting, strPackageID, m_strPkgDesc, m_strPackageType, m_status);
                                UpdateProcessEmail("True", dstAppData.DataSets[i].Fields["EmailRequest_ID"].FieldValue, Convert.ToString(iNumTries));
                                blnProceed = false;                            
                            }
                            //check package disposition and assignee status
                            else{
                                string m_strAssigneeStatus = String.Empty;
                                m_strAssigneeStatus = CheckDisposition(strPackageID, strSender);

                                //based on disposition check if the user can proceed or not
                                blnProceed = CheckAssigneeStatus(m_strAssigneeStatus, strPackageID, strSender);

                                
                                //if blnProceed is false,check to see if the user is delegate eventhough the user is not authorized to approve package as main assignee for whatever reason
                                if (!blnProceed)
                                {
                                    strSenderOrig = strSender;                                              //save orignal sender info
                                    strSender = bclRouting.GetAssigneeForDelegate(strPackageID, strSender); //and change strSender to the Delegate

                                    //if sender is delegate proceed
                                    if (!String.IsNullOrEmpty(strSender))
                                    {
                                        dstUserDetails.Fields["Assignee"].FieldValue = strSender;
                                        dstUserDetails.Fields["Comment"].FieldValue += AppendCommentsforDelegates(strAction, strSender, strSenderOrig);
                                        blnProceed = true;
                                    }
                                    else //send error email with the reason they cannot approve this package
                                    {
                                        HandleErrorReason(strEmailAddress, bclRouting, strPackageID, m_strPkgDesc, m_strPackageType, m_strAssigneeStatus);
                                        UpdateProcessEmail("True", dstAppData.DataSets[i].Fields["EmailRequest_ID"].FieldValue, Convert.ToString(iNumTries));
                                    }
                                }

                                if (blnProceed)
                                {
                                    IsPkgUpd = UpdatePackageApprovalStatus(dstUserDetails, strAction);

                                    if (IsPkgUpd == false)
                                    {
                                        SendMailToHD(strMessage, dstAppData.DataSets[i].Fields["MessageSubject"].FieldValue, strSender, strDate, "Error in UpdatePackageApprovalStatus");
                                        if (iNumTries >= iMaxTries)
                                            UpdateProcessEmail("True", dstAppData.DataSets[i].Fields["EmailRequest_ID"].FieldValue, Convert.ToString(iNumTries));
                                        else
                                            UpdateProcessEmail("False", dstAppData.DataSets[i].Fields["EmailRequest_ID"].FieldValue, Convert.ToString(iNumTries));
                                    }
                                    else
                                    {
                                        UpdateProcessEmail("True", dstAppData.DataSets[i].Fields["EmailRequest_ID"].FieldValue, Convert.ToString(iNumTries));
                                    }
                                }
                                else
                                {
                                    AddLog("CheckAssigneStatus returned false.Hence no update");                                    
                                }
                            }
                        }
                        else
                        {
                            AddLog(dstReturn.DataSets[0].Fields["ErrDesc"].FieldValue);
                            //RM - When the user doesn't have CanAccess rights, set the email as processed, send email as not an assignee
                            if (dstReturn.DataSets[0].Fields["ErrDesc"].FieldValue == "InValid user- Access Denied")
                            {
                                HandleErrorReason(strEmailAddress, bclRouting, strPackageID, "", "", RT_USERISNOTASSIGNEE);
                                UpdateProcessEmail("True", dstAppData.DataSets[i].Fields["EmailRequest_ID"].FieldValue, Convert.ToString(iNumTries));
                            }
                            else
                                SendMailToHD(strMessage, dstAppData.DataSets[i].Fields["MessageSubject"].FieldValue, strSender, strDate, strPackageID + ":" + dstReturn.DataSets[0].Fields["ErrDesc"].FieldValue);
                            

                        }
                    }
                }
            }



            catch (Exception ex)
            {
                AddLog(String.Concat(strError, strIN, MethodInfo.GetCurrentMethod().Name));
                AddLog(ex.Message);
                SendMailToHD("Routing Approval Error", "Routing Approval Error", strHDEMailID, DateTime.Now.ToString(), "Routing Approval Error");
            }

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

        }

        private static void SendMailToApprover(string strEmailAddress, Routing_bc bclRouting, string strPackageID, string m_strPkgDesc, string m_strPackageType, string errorMsg)
        {
            strEnvironment = ConfigurationManager.AppSettings["Environment"].ToString();
            
            string strErrorMessage = errorMsg;
            StringBuilder sbMailBody = new StringBuilder(4096);
            sbMailBody.Append("<html>"); //start of html and body tags
            sbMailBody.Append("<body>");
            sbMailBody.Append(String.Concat("<b>", "Email Approval Failed. Please find the below package details", "</b><br \\>"));
            sbMailBody.Append(String.Concat("<font style='color:red;'>", "Error Details: " + strErrorMessage, "</font><br \\><br \\>"));
            if (m_strPackageType.Length > 0)
            {
                sbMailBody.Append(String.Concat("<b>", STR_ROUTINGTYPE, "</b><br>"));
                sbMailBody.Append(String.Concat("<font style='color:red;'>", m_strPackageType, "</font><br><br>"));
            }
            if (m_strPkgDesc.Length > 0)
            {
                sbMailBody.Append(String.Concat("<b>", STR_ROUTINGDESC, "</b><br>"));
                sbMailBody.Append(String.Concat("<span style='word-wrap:break-word;'>" + m_strPkgDesc + "</span>", "<br><br>"));
            }
            sbMailBody.Append(String.Concat("<b>", STR_ROUTING, "</b><br>"));
            _MailAssigneeDetails(sbMailBody, strPackageID, bclRouting);

            bclRouting.SendMailToHD(sbMailBody, strEmailAddress, strTakataNetEmail);
        }

        private static void _MailAssigneeDetails(StringBuilder sbMailBody, string strPackageId, Routing_bc bclRouting)
        {
            CDataSet dstAssigneeParams = new CDataSet();
            try
            {
                CDataSet dstParams = new CDataSet();
                dstParams.Fields.Add("PackageId", strPackageId);
                dstAssigneeParams.LoadXML(bclRouting.GetApproversList(dstParams.XML), false);

                if (dstAssigneeParams.HasDataSets == true & dstAssigneeParams.Fields.Exists("RecordNotFound") == false)
                {
                    //preapring header for table
                    if ((dstAssigneeParams.GetInflatedDataSet(true).Tables[0].Rows.Count > 0))
                    {
                        sbMailBody.Append("<table style='font-family: Arial;font-size: 10pt;'><tr bgcolor=#D8D8D8>");
                        //start of assginee details table
                        sbMailBody.Append(string.Concat("<td>", STR_ASSIGNEE, "</td>"));
                        //assginee name
                        sbMailBody.Append(string.Concat("<td>", STR_DISPOSITIONDATE, "</td>"));
                        //disposittion date
                        sbMailBody.Append(string.Concat("<td>", STR_DISPOSITION, "</td>"));
                        //current dispasotion
                        sbMailBody.Append(string.Concat("<td>", STR_COMMENTS, "</td>"));
                        //current dispasotion
                        sbMailBody.Append("</tr></font>");
                    }
                    int intCounter = 0;
                    //build no of row of assginee details
                    for (intCounter = 0; intCounter <= dstAssigneeParams.DataSets.Count - 1; intCounter++)
                    {
                        var _with2 = dstAssigneeParams.DataSets[intCounter];
                        sbMailBody.Append("<tr>");
                        if (string.Compare(_with2.Fields["Actual"].FieldValue, string.Empty) == 0)
                        {
                            sbMailBody.Append(string.Concat("<td>", _with2.Fields["Assignee"].FieldValue, "</td>"));
                        }
                        else
                        {
                            sbMailBody.Append(string.Concat("<td>", _with2.Fields["Actual"].FieldValue, "</td>"));
                        }
                        if (string.Compare(_with2.Fields["DispositionDate"].FieldValue, string.Empty) == 0)
                        {
                            sbMailBody.Append("<td>&nbsp;</td>");
                        }
                        else
                        {
                            sbMailBody.Append(string.Concat("<td>", string.Concat(GetBriefDateTime(_with2.Fields["DispositionDate"].FieldValue, "", "MMM dd, yyyy hh:mm tt", false), STR_UTCDATE), "</td>"));
                        }
                        if (string.Compare(_with2.Fields["Description"].FieldValue.ToUpper(), "IN-PROCESS", true) == 0)
                        {
                            sbMailBody.Append(string.Concat("<td>", "Initiated", "</td>"));
                        }
                        else
                        {
                            sbMailBody.Append(string.Concat("<td>", _with2.Fields["Description"].FieldValue, "</td>"));
                        }
                        if (string.Compare(_with2.Fields["Comment"].FieldValue, string.Empty) == 0)
                        {
                            sbMailBody.Append("<td>&nbsp;</td>");
                        }
                        else
                        {
                            sbMailBody.Append(string.Concat("<td>", _with2.Fields["Comment"].FieldValue, "</td>"));
                        }
                        sbMailBody.Append("</tr>");
                    }
                    sbMailBody.Append("</Table>");
                    //end of assignee details table
                }
                //return string.Empty;
            }
            catch (Exception ex)
            {
                AddLog(MethodInfo.GetCurrentMethod().Name + " " + strError);
                AddLog(ex.Message);
            }
        }

        private static string GetBriefDateTime(string UTCDateTime, string UICulture, string strFormat, bool useGMT)
        {
            //Declare local variables here
            DateTime dtDateTime = default(DateTime);
            string strReturn = string.Empty;
            DateTimeFormatInfo objDateFormatInfo = new DateTimeFormatInfo();
            if (UICulture == string.Empty)
            {
                UICulture = CultureInfo.CurrentUICulture.ToString();
            }
            CultureInfo MyCultureInfo = new CultureInfo(UICulture);
            try
            {
                if ((UTCDateTime != null) && 3 < UTCDateTime.Length)
                {
                    objDateFormatInfo.ShortDatePattern = "M/d/yyyy";
                    objDateFormatInfo.LongTimePattern = "h:mm:ss tt";
                    MyCultureInfo.DateTimeFormat = objDateFormatInfo;
                    if (useGMT)
                    {
                        dtDateTime = DateTime.ParseExact(UTCDateTime, "G", MyCultureInfo);
                    }
                    else
                    {
                        dtDateTime = DateTime.ParseExact(UTCDateTime, "G", MyCultureInfo).ToLocalTime();
                    }

                    strReturn = dtDateTime.ToString(strFormat);
                }
            }
            catch
            {
            }
            return strReturn;
        }

        private static void UpdateProcessEmail(string IsProcessed, string ID, string NumTries)
        {
            CDataSet dstParams = new CDataSet();

            try
            {
                if (iCurLogLevelSet > iLogLevel)
                    AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));


                dstParams.Fields.Add("IsProcessed", IsProcessed);
                dstParams.Fields.Add("EmailRequest_ID", ID);
                dstParams.Fields.Add("NumTries", NumTries);

                Routing_bc bclRouting = new Routing_bc();
                dstParams.LoadXML(bclRouting.UpdateEmailMessage(dstParams.XML), false);


            }
            catch (Exception ex)
            {
                AddLog(MethodInfo.GetCurrentMethod().Name + " " + strError);
                AddLog(ex.Message);
            }

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

        }

        private static void SendMailToHD(string strMessage, string strSubject, string strSender, string strDate, string strErrorMessage)
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            try
            {
                StringBuilder sbMailBody = new StringBuilder(4096);
                sbMailBody.Append("<html>"); //start of html and body tags
                sbMailBody.Append("<body>");
                sbMailBody.Append(String.Concat("<b>", "Email Approval Failed. Please find the below package details", "</b><br>"));
                sbMailBody.Append(String.Concat("<font style='color:red;'>", "Comments: " + strMessage, "</font><br><br>"));
                sbMailBody.Append(String.Concat("<font style='color:red;'>", "Subject Details: " + strSubject, "</font><br><br>"));
                sbMailBody.Append(String.Concat("<font style='color:red;'>", "Sender Details: " + strSender, "</font><br><br>"));
                sbMailBody.Append(String.Concat("<font style='color:red;'>", "Date Details: " + strDate, "</font><br><br>"));
                sbMailBody.Append(String.Concat("<font style='color:red;'>", "Error Details: " + strErrorMessage, "</font><br><br>"));
                sbMailBody.Append("</body></html>");

                using (Routing_bc bclRouting = new Routing_bc())
                {

                    bclRouting.SendMailToHD(sbMailBody, strHDEMailID, strTakataNetEmail);
                }
            }
            catch (Exception ex)
            {
                AddLog(MethodInfo.GetCurrentMethod().Name + " " + strError);
                AddLog(ex.Message);
            }

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

        }

        private static string FormatComment(string strMessage)
        {
            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));

            strMessage = strMessage.Replace(strMessage1, String.Empty);
            strMessage = strMessage.Replace(strMessage2, String.Empty);
            strMessage = strMessage.Replace(strMessage3, String.Empty);
            strMessage = strMessage.Replace(Regex.Replace(strMessage4, @"\s{2,}", " "), String.Empty); //RM - 07/20/2018 - If text has multiple spaces, replace with single space
            strMessage = strMessage.Replace(strMessage5, String.Empty);

            if (iCurLogLevelSet > iLogLevel)
                AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

            return strMessage;
        }

        /// <summary>
        /// RM - 07/20/2018 - If email has comments tags this could be used to get the comments instead of FormatComment() 
        /// </summary>
        /// <param name="strMessage"></param>
        /// <returns></returns>
        //private static string FormatCommentWithTags(string strMessage)
        //{
        //    if (iCurLogLevelSet > iLogLevel)
        //        AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strStarted));
        //    try 
        //    {	        
        //        int startIndexOfComments = strMessage.ToLower().IndexOf("<comments>");
        //        int endIndexOfComments = strMessage.ToLower().IndexOf("</comments>");
        //        if(endIndexOfComments > startIndexOfComments)
        //        {
        //        int commentsLen = endIndexOfComments - startIndexOfComments;
        //        strMessage.Substring(startIndexOfComments,commentsLen);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }

        //    if (iCurLogLevelSet > iLogLevel)
        //        AddLog(String.Concat(MethodInfo.GetCurrentMethod().Name, " ", strEnded));

        //    return strMessage;
        //}

        private static string GetUserFullName(string strUserID)
        {
            string strReturnVal = string.Empty;
            CDataSet dstParams = new CDataSet();
            dstParams.Fields.Add("UserName", strUserID);
            using (Routing_bc bclRouting = new Routing_bc())
            {
                dstParams.LoadXML(bclRouting.GetUserFullName(dstParams.XML), false);
            }
            if ((!dstParams.IsErrorDataSet) && (dstParams.DataSets.Count > 0))
                strReturnVal = dstParams.DataSets[0].Fields["fullName"].FieldValue;

            if (strReturnVal == null)
                strReturnVal = strUserID;

            return strReturnVal;
        }

        private static string AppendCommentsforDelegates(String strAction, string strSender, string strSenderOrig)
        {
            String strCommentReturn = string.Empty;

            using (Routing_bc bclRouting = new Routing_bc())
            {
                if (strAction.ToUpper() == APPR_GUID)
                    strCommentReturn = String.Concat(" Approved by ", bclRouting.GetUserFullNameStub(strSenderOrig), " on behalf of ", bclRouting.GetUserFullNameStub(strSender), ".");
                else if (strAction.ToUpper() == REJ_GUID)
                    strCommentReturn = String.Concat(" Rejected by ", bclRouting.GetUserFullNameStub(strSenderOrig), " on behalf of ", bclRouting.GetUserFullNameStub(strSender), ".");
                else if (strAction.ToUpper() == CLAR_GUID)
                    strCommentReturn = String.Concat(" Seek Clarification by ", bclRouting.GetUserFullNameStub(strSenderOrig), " on behalf of ", bclRouting.GetUserFullNameStub(strSender), ".");
            }

            return strCommentReturn;

        }

        /// <summary>
        /// This function deletes the html tags from an html string.
        /// </summary>
        /// <param name="strHtml">Body of the email</param>
        /// <returns></returns>
        private static string RemoveTags(string strHtml)
        {
            int iStarteIndex;
            int iLastIndex;
            try
            {
                strHtml = strHtml.Replace("\r", String.Empty);
                strHtml = strHtml.Replace("\n", String.Empty);
                //We are interested only in parsing the html from the body tag/
                iStarteIndex = strHtml.IndexOf("<body");
                strHtml = strHtml.Substring(iStarteIndex);
                while (true)
                {
                    iStarteIndex = strHtml.IndexOf("<");
                    iLastIndex = strHtml.IndexOf(">");
                    if ((iStarteIndex < 0) || (iLastIndex < 0))
                        break;

                    strHtml = strHtml.Remove(iStarteIndex, iLastIndex - iStarteIndex + 1);
                }
            }
            catch (Exception ex)
            {
                AddLog(ex.StackTrace);
                AddLog(ex.Message);
                strHtml = string.Empty;
            }
            if (strHtml.StartsWith("<html>"))
                strHtml = string.Empty;

            strHtml = strHtml.Replace(strMessage1, String.Empty);
            strHtml = strHtml.Replace(strMessage2, String.Empty);
            strHtml = strHtml.Replace(strMessage3, String.Empty);
            strHtml = strHtml.Replace(strMessage4, String.Empty);
            strHtml = strHtml.Replace(strMessage5, String.Empty);

            return strHtml;
        }

    }
}
