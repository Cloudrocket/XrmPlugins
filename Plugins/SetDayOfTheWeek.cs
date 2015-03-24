// <copyright file="GetDayOfTheWeek.cs" company="">
// Copyright (c) 2014 All Rights Reserved
// </copyright>
// <author></author>
// <date>1/6/2014 2:35:32 PM</date>
// <summary>Implements the GetDayOfTheWeek Workflow Activity.</summary>
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

    public sealed class SetDayOfTheWeek : CodeActivity
    {
        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered SetDayOfTheWeek.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("SetDayOfTheWeek.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // Get the entity reference from an input argument.
                // EntityReference inputEntity = this.inputPageView.Get(executionContext);

                // Get the entity reference from the workflow context.
                Entity inputEntity = service.Retrieve(
                       context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet { AllColumns = true });

                // Build the entity request and get the entity.
                RetrieveRequest entityRequest = new RetrieveRequest
                {
                    ColumnSet = new ColumnSet { AllColumns = true, },
                    Target = new EntityReference
                    {
                        Id = inputEntity.Id,
                        LogicalName = inputEntity.LogicalName,
                    }
                };

                Entity entity = (Entity)((RetrieveResponse)service.Execute(entityRequest)).Entity;

                #region Set the day of the week, in Pacific Time Zone from UTC cldrkt_createdon.

                DateTime createdOn = (DateTime)entity["createdon"];
                createdOn = DateTime.SpecifyKind(createdOn, DateTimeKind.Utc);

                TimeZoneInfo timeZone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                string dayOfWeek = TimeZoneInfo.ConvertTimeFromUtc(createdOn, timeZone).DayOfWeek.ToString().Substring(0, 3);

                // Get the Days of the Week Option Set metadata.
                RetrieveOptionSetRequest optionSetRequest = new RetrieveOptionSetRequest { Name = "cldrkt_daysoftheweek" };
                RetrieveOptionSetResponse optionSetResponse = (RetrieveOptionSetResponse)service.Execute(optionSetRequest);

                OptionSetMetadata optionSetMetaData = (OptionSetMetadata)optionSetResponse.OptionSetMetadata;

                // Look up the OptionSetValue Value using dayOfWeek.
                OptionSetValue optionSetValue = new OptionSetValue
                {
                    Value = (Int32)optionSetMetaData.Options
                    .FirstOrDefault(o => o.Label.UserLocalizedLabel.Label == dayOfWeek).Value,
                };

                // outputDayOfTheWeek.Set(executionContext, dayOfWeek);

                entity["cldrkt_createdonweekdayoptionset"] = optionSetValue;

                #endregion Set the day of the week, in Pacific Time Zone from UTC cldrkt_createdon.

                service.Update(entity);
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting SetDayOfTheWeek.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        // Define Input/Output Arguments
        //[RequiredArgument]
        //[Input("Web Page View")]
        //[ReferenceTarget("cldrkt_pageview")]
        //public InArgument<EntityReference> inputPageView { get; set; }

        // Use this section to set up output parameters to set on workflow form.
        //[Output("OptionSet Day Of The Week output")]
        //[AttributeTarget("cldrkt_pageview", "cldrkt_createdonweekdayoptionset")]
        // public OutArgument<OptionSetValue> outputOptionSetValueDayOfTheWeek { get; set; }

        [Output("DayOfWeek")]
        public OutArgument<string> outputDayOfTheWeek { get; set; }
    }
}