// <copyright file="SetContactProperties.cs" company="">Copyright (c) 2014 All Rights Reserved</copyright>
// <author></author>
// <date>4/24/2014 12:47:17 PM</date>
// <summary>
// Implements the SetContactProperties Workflow Activity.
// </summary>
namespace Cloudrocket.Crm.Plugins
{
    using Cloudrocket.Xrm;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Workflow;
    using System.Activities;
    using System.ServiceModel;

    public sealed class SetContactProperties : CodeActivity
    {   // Use this section to set up output parameters to set on workflow form.
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

            tracingService.Trace("Entered SetContactProperties.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("SetContactProperties.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // TODO: Implement your custom Workflow business logic.

                Contact entity = (Contact)service.Retrieve(context.PrimaryEntityName,
                    context.PrimaryEntityId, new ColumnSet { AllColumns = true });

                if (entity.LogicalName != "contact")
                {
                    t.Trace("This entity isn't a Contact.");
                    throw new InvalidPluginExecutionException();
                }

                // Set the email domain (used for import correlation).

                //Helpers helper = new Helpers();

                //t.Trace("1. Checking the email address.");

                //if (helper.IsValidEmail(entity.EMailAddress1))
                //{
                //    entity.cldrkt_EmailDomain = new MailAddress(entity.EMailAddress1).Host.ToString();
                //    t.Trace("Set the email domain to " + entity.EMailAddress1);
                //}

                //Set the time zone based on the telephone area code.
                //t.Trace("3. Get the time zone.");
                //string timeZone = string.Empty;

                //t.Trace("4a. The current telephone number is " + entity.Telephone1.ToString());

                //timeZone = Helpers.GetTimeZone(entity.Telephone1);

                //t.Trace("The time zone is " + timeZone);

                //if (timeZone != "Unknown")
                //{
                //    t.Trace("4b. Getting the time zone option set.");

                //    RetrieveOptionSetRequest optionSetRequest = new RetrieveOptionSetRequest { Name = "cldrkt_timezones" };
                //    RetrieveOptionSetResponse optionSetResponse = (RetrieveOptionSetResponse)service.Execute(optionSetRequest);

                //    OptionSetMetadata optionSetMetaData = (OptionSetMetadata)optionSetResponse.OptionSetMetadata;

                //    // Look up the OptionSetValue Value using dayOfWeek.
                //    OptionSetValue optionSetValue = new OptionSetValue {
                //        Value = (Int32)optionSetMetaData.Options
                //        .FirstOrDefault(o => o.Label.UserLocalizedLabel.Label == timeZone).Value,
                //    };

                //    if (optionSetValue.Value >= 0)
                //    {
                //        t.Trace("4a. The option set value is " + optionSetValue.ToString());

                //        entity.cldrkt_TimeZone = optionSetValue;

                //        t.Trace("The entity time zone is set to " + entity.cldrkt_TimeZone.ToString());
                //    }
                //}

                // Sets cldrkt_MarketingListSeparator to allow A/B split testing with marketing lists.

                t.Trace("Split is " + SplitNumber.ToString());

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

                t.Trace("Splitter is " + entity.cldrkt_MarketingListSeparator.ToString());

                if (entity.Attributes["cldrkt_entityidstring"] != null)
                {
                    entity.Attributes["cldrkt_entityidstring"] = entity.Id.ToString();

                    t.Trace("About to update the entity.");
                    service.Update(entity);
                }
                else
                {
                    t.Trace("The Marketing Split field is not in this entity.");
                }

                //throw new InvalidPluginExecutionException("Finished processing the Contact update.");
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting SetContactProperties.Execute(), Correlation Id: {0}", context.CorrelationId);
        }
    }
}