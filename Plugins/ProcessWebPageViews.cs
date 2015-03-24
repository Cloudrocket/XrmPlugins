// <copyright file="GetCampaignResponseFromWebPageView.cs" company="Cloudrocket LLC">
// Copyright (c) 2014 All Rights Reserved
// </copyright>
// <author></author>
// <date>1/29/2014 11:48:21 AM</date>
// <summary>
// Implements the GetCampaignResponseFromWebPageView Workflow Activity.
// </summary>
namespace Cloudrocket.Crm.Plugins
{
    using Cloudrocket.Xrm;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Workflow;
    using System;
    using System.Activities;
    using System.Linq;
    using System.ServiceModel;

    public sealed class ProcessWebPageViews : CodeActivity
    {
        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext) {

            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered GetCampaignResponseFromWebPageView.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("GetCampaignResponseFromWebPageView.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            ITracingService t = tracingService;

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // If we have a CampaignActivityId in the URL, create a CampaignResponse and map it
                // to the Web Page View. If we have a CampaignResponseCustomer in the URL, add that
                // Contact to the CampaignResponse and Web Page View.

                // Get the Page View from the workflow context.
                t.Trace("1. Get the Page View from the workflow context and default the driving variables.");
                cldrkt_pageview pageView = (cldrkt_pageview)service.Retrieve(
                    context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet { AllColumns = true });

                CampaignActivity campaignActivity = new CampaignActivity();
                CampaignResponse campaignResponse = new CampaignResponse();
                Contact contact = new Contact();

                #region Process the Campaign Customer, if any.

                // Get the Campaign Customer, if any.
                t.Trace("c1. Get the Campaign Customer, if any...");

                Guid contactId = Guid.Empty;

                t.Trace("pageView.cldrkt_CampaignResponseCustomerId: " + pageView.cldrkt_CampaignResponseCustomerId);
                if (String.IsNullOrWhiteSpace(pageView.cldrkt_CampaignResponseCustomerId))
                {
                    contactId = Guid.Empty;
                    t.Trace("contactId: " + contactId.ToString());
                }
                else
                {
                    Guid.TryParse(pageView.cldrkt_CampaignResponseCustomerId, out contactId);

                    if (contactId != Guid.Empty)
                    {
                        t.Trace("c2. Look up the  Campaign Customer...");

                        contact = (Contact)service.Retrieve(Contact.EntityLogicalName, contactId, new ColumnSet { AllColumns = true });
                        t.Trace("contactId: " + contactId.ToString());
                        t.Trace("contact.Id: " + contact.Id.ToString());

                        if (contact != null)
                        {
                            // Add the Campaign Activity Customer to the Campaign Response and Page View
                            t.Trace("c3. Add the Campaign Activity Customer to the Campaign Response and Page View");

                            campaignResponse.Customer = new ActivityParty[]
                        {
                            new ActivityParty {PartyId = new EntityReference (contact.LogicalName, contact.Id)}
                        };
                            pageView.cldrkt_CampaignResponseCustomer = new EntityReference {
                                Id = contact.Id,
                                LogicalName = contact.LogicalName,
                            };
                        }
                    }
                }

                #endregion Process the Campaign Customer, if any.

                #region Process the Campaign Activity, if any.

                // Get the Campaign Activity, if any.
                t.Trace("ca1. Get the Campaign Activity, if any...");

                Guid campaignActivityId = Guid.Empty;

                t.Trace("pageView.cldrkt_CampaignActivityId: " + pageView.cldrkt_CampaignActivityId);
                if (String.IsNullOrWhiteSpace(pageView.cldrkt_CampaignActivityId))
                {
                    campaignActivityId = Guid.Empty;
                    t.Trace("campaignActivityId: " + campaignActivityId);
                }
                else
                {
                    Guid.TryParse(pageView.cldrkt_CampaignActivityId, out campaignActivityId);

                    if (campaignActivityId != Guid.Empty) // Look up the Campaign Activity
                    {
                        t.Trace("ca2. Look up the Campaign Activity...");

                        campaignActivity = (CampaignActivity)service.Retrieve(
                            CampaignActivity.EntityLogicalName, campaignActivityId, new ColumnSet { AllColumns = true });
                        t.Trace("campaignActivityId: " + campaignActivityId);
                        t.Trace("campaignActivity.Id: " + campaignActivity.Id.ToString());

                        if (campaignActivity != null) // Process for a Campaign Activity
                        {
                            // Create a Campaign Response.
                            t.Trace("ca3. Create a Campaign Response...");

                            campaignResponse.ChannelTypeCode = new OptionSetValue(636280000);

                            campaignResponse.OriginatingActivityId = new EntityReference {
                                Id = campaignActivity.Id,
                                LogicalName = CampaignActivity.EntityLogicalName,
                            };
                            campaignResponse.RegardingObjectId = new EntityReference // Required, must be the parent campaign
                            {
                                Id = campaignActivity.RegardingObjectId.Id,
                                LogicalName = Campaign.EntityLogicalName,
                            };
                            campaignResponse.ReceivedOn = pageView.CreatedOn;
                            campaignResponse.Subject = pageView.cldrkt_name;

                            campaignResponse.ActivityId = service.Create(campaignResponse);
                            t.Trace("campaignResponse.ActivityId: " + campaignResponse.ActivityId);
                            t.Trace("campaignResponse.Id: " + campaignResponse.Id.ToString());

                            // Update the Campaign Response.
                            t.Trace("ca4. Update the Campaign Response.");
                            t.Trace("campaignResponse.Id: " + campaignResponse.Id);

                            if (campaignResponse.Id != Guid.Empty)
                            {
                                service.Update(campaignResponse);
                                t.Trace("campaignResponse.Id = " + campaignResponse.Id.ToString());
                            }

                            // Add the Campaign Activity to the Page View.
                            t.Trace("4. Add the Campaign Activity to the Page View");

                            pageView.cldrkt_Campaign = new EntityReference {
                                Id = campaignActivity.RegardingObjectId.Id,
                                LogicalName = campaignActivity.RegardingObjectId.LogicalName,
                            };
                            pageView.cldrkt_CampaignActivity = new EntityReference {
                                Id = campaignActivity.Id,
                                LogicalName = campaignActivity.LogicalName,
                            };
                            pageView.cldrkt_CampaignResponse = new EntityReference {
                                Id = campaignResponse.Id,
                                LogicalName = campaignResponse.LogicalName,
                            };
                        }
                    }
                }

                #endregion Process the Campaign Activity, if any.

                #region Set the day of the week, in Pacific Time Zone from UTC cldrkt_createdon.

                DateTime createdOn = (DateTime)pageView["createdon"];
                createdOn = DateTime.SpecifyKind(createdOn, DateTimeKind.Utc);

                TimeZoneInfo timeZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                string dayOfWeek = TimeZoneInfo.ConvertTimeFromUtc(createdOn, timeZone).DayOfWeek.ToString().Substring(0, 3);

                // Get the Days of the Week Option Set metadata.
                RetrieveOptionSetRequest optionSetRequest = new RetrieveOptionSetRequest { Name = "cldrkt_daysoftheweek" };
                RetrieveOptionSetResponse optionSetResponse = (RetrieveOptionSetResponse)service.Execute(optionSetRequest);

                OptionSetMetadata optionSetMetaData = (OptionSetMetadata)optionSetResponse.OptionSetMetadata;

                // Look up the OptionSetValue Value using dayOfWeek.
                OptionSetValue optionSetValue = new OptionSetValue {
                    Value = (Int32)optionSetMetaData.Options
                    .FirstOrDefault(o => o.Label.UserLocalizedLabel.Label == dayOfWeek).Value,
                };

                // outputDayOfTheWeek.Set(executionContext, dayOfWeek);

                pageView["cldrkt_createdonweekdayoptionset"] = optionSetValue;

                #endregion Set the day of the week, in Pacific Time Zone from UTC cldrkt_createdon.



                // Update the Page View.
                t.Trace("10. Update the Page View.");

                service.Update(pageView);

                // throw new InvalidPluginExecutionException("Finished processing the Page View update.");
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting GetCampaignResponseFromWebPageView.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        // Define Input/Output Arguments [RequiredArgument] [Input("Web Page View")]
        // [ReferenceTarget("cldrkt_pageview")] public InArgument<EntityReference> inputPageView {
        // get; set; }

        // Use this section to set up output parameters to set on workflow form.
        //[Output("OptionSet Day Of The Week output")]
        //[AttributeTarget("cldrkt_pageview", "cldrkt_createdonweekdayoptionset")]
        // public OutArgument<OptionSetValue> outputOptionSetValueDayOfTheWeek { get; set; }

        // [Output("DayOfWeek")] public OutArgument<string> outputDayOfTheWeek { get; set; }
    }
}