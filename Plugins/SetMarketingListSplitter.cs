// <copyright file="SetMarketingListSeparator.cs" company="">
// Copyright (c) 2014 All Rights Reserved
// </copyright>
// <author></author>
// <date>4/1/2014 6:46:58 PM</date>
// <summary>
// Implements the SetMarketingListSeparator Workflow Activity.
// </summary>
namespace Cloudrocket.Crm.Plugins
{
    using Cloudrocket.Xrm;
    using Microsoft.Crm.Sdk;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Workflow;
    using System;
    using System.Activities;
    using System.Linq;
    using System.ServiceModel;

    public sealed class SetMarketingListSplitter : CodeActivity
    {
        // Use this section to set up output parameters to set on workflow form.
        [Input("Number of splits for this list")]
        // [Default("2")]
        public InArgument<int> SplitNumber { get; set; }

        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext) {

            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            ITracingService t = tracingService;

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered SetMarketingListSeparator.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("SetMarketingListSeparator.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // TODO: Implement your custom Workflow business logic.

                // Sets cldrkt_MarketingListSeparator to allow A/B split testing with marketing lists.

                t.Trace("Get the entity. ");

                Contact entity = (Contact)service.Retrieve(context.PrimaryEntityName,
                context.PrimaryEntityId, new ColumnSet { AllColumns = true });

                if (entity.LogicalName != "contact")
                {
                    t.Trace("This entity isn't a Contact.");
                    throw new InvalidPluginExecutionException();
                }

                int i = 0;
                int split = SplitNumber.Get(executionContext);
                string s = null;

                byte[] guid = entity.ContactId.Value.ToByteArray();

                foreach (byte b in guid) { i = i + (int)b; }
                i = i % split;

                switch (i)
                {
                    case 0: s = "A"; break;
                    case 1: s = "B"; break;
                    case 2: s = "C"; break;
                    case 3: s = "D"; break;
                    case 4: s = "E"; break;
                    case 5: s = "F"; break;
                    case 6: s = "G"; break;
                    case 7: s = "H"; break;
                    case 8: s = "I"; break;
                    case 9: s = "J"; break;
                    case 10: s = "K"; break;
                    case 11: s = "L"; break;
                    case 12: s = "M"; break;
                    case 13: s = "N"; break;
                    case 14: s = "O"; break;
                    case 15: s = "P"; break;
                    case 16: s = "A"; break;
                    case 17: s = "Q"; break;
                    case 18: s = "R"; break;
                    case 19: s = "S"; break;
                    case 20: s = "T"; break;
                    case 21: s = "U"; break;
                    case 22: s = "V"; break;
                    case 23: s = "W"; break;
                    default: s = "Undefined"; break;
                }

                entity.cldrkt_MarketingListSeparator = s;

                service.Update(entity);
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting SetMarketingListSplitter.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        //[Output("OptionSet Day Of The Week output")]
        //[AttributeTarget("cldrkt_pageview", "cldrkt_createdonweekdayoptionset")]
        // public OutArgument<OptionSetValue> outputOptionSetValueDayOfTheWeek { get; set; }

        // [Output("DayOfWeek")] public OutArgument<string> outputDayOfTheWeek { get; set; }
    }
}