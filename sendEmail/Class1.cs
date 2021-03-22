using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;


namespace sendEmail
{
    public class Class1 : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                if (context.Depth > 2) {
                    try
                    {
                        Entity entity = (Entity)context.InputParameters["Target"];
                        Entity BE = service.Retrieve("crec2_business", entity.Id, new ColumnSet(true));


                        String subject = BE.GetAttributeValue<String>("crec2_subject");
                        String emailBody = BE.GetAttributeValue<String>("new_emailbody");
                        EntityReference Cc = BE.GetAttributeValue<EntityReference>("new_mailcc");
                        EntityReference Bcc = BE.GetAttributeValue<EntityReference>("new_mailbcc");
                        EntityReference to = BE.GetAttributeValue<EntityReference>("new_mailto");
                        String[] emailsyntax = emailBody.Split('\n');
                        String email = "<html><body><div>";
                        foreach (String line in emailsyntax)
                        {
                            tracing.Trace("email" + email);
                            email = email + line + "<br>";
                        }
                        email = email + "</div></body></html>";

                        Entity ccActivityParty = new Entity("activityparty");
                        Entity bccActivityParty = new Entity("activityparty");
                        Entity toActivityParty = new Entity("activityparty");


                        ccActivityParty["partyid"] = Cc;
                        bccActivityParty["partyid"] = Bcc;
                        toActivityParty["partyid"] = to;



                        Entity entEmail = new Entity("email");
                        entEmail["subject"] = subject;
                        entEmail["description"] = email;
                        //entEmail["from"] = "Saurabh.tripathi@online24x7.net";
                        entEmail["to"] = new Entity[] { toActivityParty };
                        entEmail["cc"] = new Entity[] { ccActivityParty };
                        entEmail["bcc"] = new Entity[] { bccActivityParty };

                        entEmail["regardingobjectid"] = new EntityReference("crec2_business", BE.Id);
                        Guid emailId = service.Create(entEmail);

                        //Send email
                        SendEmailRequest sendEmailReq = new SendEmailRequest()
                        {
                            EmailId = emailId,
                            IssueSend = true
                        };
                        SendEmailResponse sendEmailRes = (SendEmailResponse)service.Execute(sendEmailReq);

                    }
                    catch (Exception e)
                    {
                        throw new InvalidPluginExecutionException("error : " + e.Message);
                    }
                }

            }
        }
    }
}
