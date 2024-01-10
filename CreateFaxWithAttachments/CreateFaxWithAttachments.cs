using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using System.Activities;
using Microsoft.SharePoint.Client;
using System.IO;
using System.Net;
using System.Text;

namespace CreateFaxWithAttachments
{
    public class CreateFaxWithAttachments : CodeActivity
    {
         [Input("DistributionLog")]
        [ReferenceTarget("pps_distributions")]
        public InArgument<EntityReference> distributionsReference { get; set; }

        [Output("FaxIDFromPlugin")]
        public OutArgument<string> faxIDReturn { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            try
            {

                tracingService.Trace("Create Email With Attachments Plugin Started");
                IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

                IOrganizationServiceFactory serviceFactoryProxy = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService serviceProxy = serviceFactoryProxy.CreateOrganizationService(context.UserId);

                EntityReference distributionDocumentsReference = distributionsReference.Get<EntityReference>(executionContext);
                QueryExpression distributionDocumentsQuery = new QueryExpression
                {
                    EntityName = pps_distributiondocuments.EntityLogicalName,
                    ColumnSet = new ColumnSet("pps_sharepointabsoluteurl", "pps_name", "pps_mimetype", "statuscode"),
                    Criteria = new FilterExpression
                    {
                        Conditions = 
                        {
                            new ConditionExpression 
                            {
                                AttributeName = "pps_regardingid",
                                Operator = ConditionOperator.Equal,
                                Values = { distributionDocumentsReference.Id }
                            }
                        }
                    }
                };
                DataCollection<Entity> distributionDocs = serviceProxy.RetrieveMultiple(distributionDocumentsQuery).Entities;

                pps_distributions distributionLogsEntity = new pps_distributions();
                ColumnSet distributionLogAttributes = new ColumnSet(new string[] { "pps_distributioneventid", "pps_lastname", "pps_firstname", "pps_participantfax" });
                distributionLogsEntity = (pps_distributions)serviceProxy.Retrieve(pps_distributions.EntityLogicalName, distributionDocumentsReference.Id, distributionLogAttributes);

                //Create Email
                Entity newEmail = new Entity("email");

                Entity fromParty = new Entity("activityparty");
                //fromParty["partyid"] = new EntityReference("systemuser", context.UserId);
                //TEST: User for the account "403b Distributions" 
                fromParty["partyid"] = new EntityReference("systemuser", Guid.Parse("A96D2AB1-C3D6-E711-80E5-005056941AF2"));
                //PROD: User for the account "403b Distributions" 
                //fromParty["partyid"] = new EntityReference("systemuser", Guid.Parse("AEF41EBC-29D9-E711-80E6-005056941AF2"));

                //Set email or Fax type
                OptionSetValue faxTypeSet = new OptionSetValue();
                faxTypeSet.Value = 927830001;
                distributionLogsEntity.Attributes["pps_type"] = faxTypeSet;
                tracingService.Trace(distributionLogsEntity.pps_participantfax);
                newEmail["pps_faxnumber"] = distributionLogsEntity.pps_participantfax;
                newEmail["pps_type"] = faxTypeSet;
                newEmail["from"] = new Entity[] { fromParty };

                Entity toParty = new Entity("activityparty");
                toParty["participationtypemask"] = new OptionSetValue(0);
                if (distributionLogsEntity.pps_participantfax != null && distributionLogsEntity.pps_participantfax != string.Empty)
                {
                    toParty["addressused"] = distributionLogsEntity.pps_participantfax + "@fax.penserv.com";
                    toParty["fullname"] = distributionLogsEntity.pps_firstname + " " + distributionLogsEntity.pps_lastname;
                    newEmail["to"] = new Entity[] { toParty };
                }

                newEmail["subject"] = "FAX: " + distributionLogsEntity.pps_distributioneventid.ToString();
                newEmail["description"] = distributionLogsEntity.pps_firstname + " " + distributionLogsEntity.pps_lastname;
                newEmail["regardingobjectid"] = new EntityReference(pps_distributions.EntityLogicalName, distributionDocumentsReference.Id);
                Guid emailId = serviceProxy.Create(newEmail);
                tracingService.Trace("Fax Created");

                foreach (Entity entity in distributionDocs)
                {
                    try
                    {
                        pps_distributiondocuments ddEntity = (pps_distributiondocuments)entity;
                        if (ddEntity.pps_SharePointAbsoluteURL.ToString() != null && ddEntity.pps_SharePointAbsoluteURL.ToString() != string.Empty && ddEntity.statuscode.Value.ToString() == "1")
                        {
                            tracingService.Trace(ddEntity.pps_SharePointAbsoluteURL.ToString());
                            Stream attachmentStream = DownloadSharePointFile(ddEntity.pps_SharePointAbsoluteURL.ToString(), tracingService);

                            //Get file extension
                            string fileExtension = ddEntity.pps_SharePointAbsoluteURL.ToString().Substring(ddEntity.pps_SharePointAbsoluteURL.ToString().LastIndexOf("."));
                            tracingService.Trace(fileExtension);

                            Entity attachment = new Entity("activitymimeattachment");
                            attachment["subject"] = "Fax Attachment";
                            string fileName = ddEntity.pps_name.ToString() + fileExtension;
                            attachment["filename"] = fileName;

                            byte[] buf = ToByteArray(attachmentStream);

                            attachment["body"] = Convert.ToBase64String(buf);
                            attachment["mimetype"] = ddEntity.pps_MimeType.ToString();
                            //attachment["mimetype"] = "application/pdf";
                            attachment["attachmentnumber"] = 1;
                            attachment["objectid"] = new EntityReference("email", emailId);
                            attachment["objecttypecode"] = Email.EntityLogicalName;
                            serviceProxy.Create(attachment);
                            tracingService.Trace("Attachment Created");
                        }
                    }
                    catch (Exception ex)
                    {
                        tracingService.Trace(ex.Message);
                    }
                }
                tracingService.Trace(emailId.ToString());
                faxIDReturn.Set(executionContext, emailId.ToString());
            }
            catch (Exception ex)
            {
                tracingService.Trace(ex.Message);
            }
        }

        public Stream DownloadSharePointFile(string sharePointURL, ITracingService tracingService)
        {
            try
            {
                WebRequest request = WebRequest.Create(new Uri(sharePointURL, UriKind.Absolute));
                request.Credentials = new NetworkCredential("test", "PenServ1");
                WebResponse response = request.GetResponse();
                Stream fs = response.GetResponseStream() as Stream;
                return fs;
            }
            catch (Exception ex)
            {
                tracingService.Trace(ex.Message);
                return null;
            }
        }

        public static byte[] ToByteArray(Stream stream)
        {
            using (stream)
            {
                using (MemoryStream memStream = new MemoryStream())
                {
                    stream.CopyTo(memStream);
                    return memStream.ToArray();
                }
            }
        }
    }
}