// <copyright file="DistributeCampaignActivityEmails.cs" company="">
// Copyright (c) 2014 All Rights Reserved
// </copyright>
// <author></author>
// <date>3/28/2014 9:57:34 PM</date>
// <summary>
// Implements the DistributeCampaignActivityEmails Workflow Activity.
// </summary>
namespace Cloudrocket.Crm.Plugins
{
    using System;
    using System.Activities;
    using System.Linq;
    using System.ServiceModel;
    using Cloudrocket.Xrm;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Crm.Sdk.Messages; 
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Workflow;

    public sealed class DistributeCampaignActivityEmails : CodeActivity
    {
        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            ITracingService t = tracingService;

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered DistributeCampaignActivityEmails.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("DistributeCampaignActivityEmails.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            t.Trace("Get the inputEntity. ");

            Entity inputEntity = service.Retrieve(
                    context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet { AllColumns = true });

            t.Trace("Cast to a CampaignActivity. ");

            CampaignActivity entity = (CampaignActivity)inputEntity;

            try
            {
                // This custom workflow runs a DistributeCampaignActivityRequest using the email
                // template Guid set in the Campaign Activity.

                # region Campaign Activity entity error checking.
                t.Trace("Error checking. ");

                // Make sure this entity is a Campaign Activity, its Status Reason is Open, and it
                // has a valid Email Template.

                if (inputEntity.LogicalName != "campaignactivity")
                {
                    t.Trace("This isn't a Campaign Activity.");
                    throw new InvalidPluginExecutionException();
                }

                if (String.IsNullOrEmpty(entity.cldrkt_EmailTemplateID))
                {
                    t.Trace("The email template field is empty.");
                    throw new InvalidPluginExecutionException();
                }

                Guid emailTemplateId = Guid.Empty;

                Guid.TryParse(entity.cldrkt_EmailTemplateID, out emailTemplateId); // Guid.Empty if false.

                if (emailTemplateId == Guid.Empty)
                {
                    t.Trace("The email template field has data, but it isn't a valid Guid.");
                    throw new InvalidPluginExecutionException();
                }

                QueryExpression templateQuery = new QueryExpression
                {
                    EntityName = Template.EntityLogicalName,
                    ColumnSet = new ColumnSet { AllColumns = true },
                    Criteria = new FilterExpression
                    {
                        Conditions = {
                            new ConditionExpression {
                                AttributeName = "templateid",
                                Operator = ConditionOperator.Equal,
                                Values = { emailTemplateId},
                            }
                        }
                    }
                };

                t.Trace("Looking up the email Template.");
                Entity template = service.RetrieveMultiple(templateQuery).Entities.FirstOrDefault();

                if (template == null)
                {
                    t.Trace("This is a valid Guid, but it doesn't match an existing email template.");
                    throw new InvalidPluginExecutionException();
                }

                # endregion

                // Create a DistributeCampaignActivity message using the email template set in the CA.

                t.Trace("Starting the DistributeCampaignActivityRequest. ");
                t.Trace("Owning  User is: " + entity.OwningUser.Id);

                // Set the campaign email sender.
                ActivityParty[] from = new ActivityParty[]{
                    new ActivityParty {
                        PartyId = new EntityReference { Id = entity.OwningUser.Id, LogicalName = SystemUser.EntityLogicalName },
                    }
                };

                t.Trace("Sender Id is " + from[0].Id.ToString());

                DistributeCampaignActivityRequest request = new DistributeCampaignActivityRequest
                {
                    // These properties are all required:
                    Activity = new Email { From = from },
                    CampaignActivityId = entity.Id,
                    Owner = new EntityReference
                    {
                        // Id = context.UserId,  // was context.Id
                        Id = entity.OwningUser.Id,
                        LogicalName = SystemUser.EntityLogicalName,
                    },
                    OwnershipOptions = PropagationOwnershipOptions.ListMemberOwner,
                    PostWorkflowEvent = true, // True if this is an asynchronous job, false if using Word mail merge.
                    Propagate = true, // True to create, send, and mark emails complete.  False to create but not send.
                    SendEmail = true, // True to send emails automatically, false to not send emails.
                    TemplateId = Guid.Parse(entity.cldrkt_EmailTemplateID),
                };

                t.Trace("The request owner is " + request.Owner.ToString());

                t.Trace("Executing the DistributeCampaignActivityRequest. ");
                service.Execute(request);

                //throw new InvalidPluginExecutionException("Finished processing the scheduled Campaign Activity send email.");
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                t.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting DistributeCampaignActivityEmails.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        // Define Input/Output Arguments [RequiredArgument] [Input("Campaign Activity")]
        // [ReferenceTarget("systemuser")]
        // public InArgument<EntityReference> inputSystemUser { get; set; }

        // Use this section to set up output parameters to set on workflow form.
        //[Output("OptionSet Day Of The Week output")]
        //[AttributeTarget("cldrkt_pageview", "cldrkt_createdonweekdayoptionset")]
        // public OutArgument<OptionSetValue> outputOptionSetValueDayOfTheWeek { get; set; }
    }
}