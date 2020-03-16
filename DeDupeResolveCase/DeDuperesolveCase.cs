//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;

//namespace EnquiryAutoResolve
//{
//    public class Class1
//    {
//    }
//}


/*using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoRespond
{
    class AutoRespond
    {
    }
}*/


using System;
using System.ServiceModel;
using Microsoft.Crm.Sdk.Messages;

// Microsoft Dynamics CRM namespace(s)
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;

namespace DeDupeResolveCase.Plugins
{
    public class DeDuperesolveCase : IPlugin
    {
        /// <summary>
        /// A plug-in that creates a case, queue item and auto-response email record when a new, previously untracked incoming email is created.
        /// Thomas Adams, University of Bath, 29/10/19
        /// </summary>
        /// <remarks>Register this plug-in on the Create message, email entity,
        /// and asynchronous mode.
        /// </remarks>
        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService =
                (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                Entity entity = (Entity)context.InputParameters["Target"];

                string strEmailDirection, strEmailSubject, strTo, strFrom, strEmailTo, strEmailFrom, strModifiedBy;
                Boolean blnContactFound = false;
                Boolean blnCloseForDeDupe = false;
                //EntityReference contactID = new EntityReference("contact");
                Guid contactID = new Guid();
                Guid ownerID = new Guid();
                Guid queueID = new Guid();
                Guid caseID = new Guid();
                Guid queueItemID = new Guid();
                Guid emailID = new Guid();
                Guid templateEmailID = new Guid();
                Guid regardingObjectID = new Guid();
                Guid incidentID = new Guid();

                string strCategory = "";

                // Verify that the target entity represents an account.
                // If not, this plug-in was not registered correctly.
                if (entity.LogicalName != "incident")
                    return;

                try
                {
                    tracingService.Trace("DeDupeResolveCase starting......");

                    //strEmailTo = "";
                    //strEmailFrom = "";

                    if (context.PostEntityImages.Contains("PostImage") &&
                    context.PostEntityImages["PostImage"] is Entity)
                    {
                        Entity postMessageImage = (Entity)context.PostEntityImages["PostImage"];
                        ////strEmailDirection = postMessageImage.FormattedValues["directioncode"];
                        ////strEmailSubject = postMessageImage.Attributes["subject"].ToString();
                        ////strTo = postMessageImage.Attributes["to"].ToString();
                        ////strFrom = postMessageImage.Attributes["from"].ToString();
                        ////emailID = (Guid)postMessageImage.Attributes["activityid"];
                        ///
                        //regardingObjectID = ((EntityReference)(postMessageImage.Attributes["regardingobjectid"])).Id;
                        incidentID = (Guid)postMessageImage.Attributes["incidentid"];

                        strModifiedBy = postMessageImage.FormattedValues["modifiedby"];
                        blnCloseForDeDupe = (Boolean)postMessageImage.Attributes["dedupe_closefordedupe"];

                        tracingService.Trace("Modified by: " + strModifiedBy);
                        tracingService.Trace("Incident ID: " + incidentID.ToString());

                        //Connect to the org service
                        // Obtain the organization service reference.
                        IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                        IOrganizationService svc = serviceFactory.CreateOrganizationService(context.UserId);

                        if (blnCloseForDeDupe == true)
                        {
                                tracingService.Trace("Going to close case...");
                                //if (strState == "Active")
                                {
                                    //logfile.WriteLine("Resolving case...");
                                    Entity IncidentResolution = new Entity("incidentresolution");
                                    IncidentResolution.Attributes["subject"] = "DeDupe auto-resolve";
                                    //IncidentResolution.Attributes["incidentid"] = new EntityReference("incident", incidentID);
                                    
                                    // Create the request to close the incident, and set its resolution to the
                                    // resolution created above
                                    CloseIncidentRequest closeRequest = new CloseIncidentRequest();
                                    closeRequest.IncidentResolution = IncidentResolution;

                                    // Set the requested new status for the closed Incident
                                    closeRequest.Status = new OptionSetValue(5);

                                    // Execute the close request
                                    CloseIncidentResponse closeResponse = (CloseIncidentResponse)svc.Execute(closeRequest);
                                }
                            }
                        }
                    else
                    {
                        tracingService.Trace("Modified not found");
                    }
                }

                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the DeDupeResolveCase plug-in.", ex);
                }

                catch (Exception ex)
                {
                    tracingService.Trace("DeDupeResolveCase Plugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}


