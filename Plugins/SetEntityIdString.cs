// <copyright file="SetEntityStringGuid.cs" company="">Copyright (c) 2014 All Rights Reserved</copyright>
// <author></author>
// <date>1/26/2014 5:29:17 PM</date>
// <summary>
// Implements the SetEntityStringGuid Workflow Activity.
// </summary>
namespace Cloudrocket.Crm.Plugins {

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
    using System.ServiceModel;

    public sealed class SetEntityIdString : CodeActivity {

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

            tracingService.Trace("Entered SetEntityStringGuid.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("SetEntityStringGuid.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // Get the entity reference from an input argument. EntityReference inputEntity = this.inputPageView.Get(executionContext);

                // Get the entity reference from the workflow context.
                Entity inputEntity = service.Retrieve(
                    context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet { AllColumns = true });

                // AllColumns doesn't return Attributes with null values. Restrict this workflow to
                // entities that have this attribute.
                inputEntity.Attributes["cldrkt_entityidstring"] = inputEntity.Id.ToString();

                service.Update(inputEntity);
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting SetEntityStringGuid.Execute(), Correlation Id: {0}", context.CorrelationId);
        }
    }
}