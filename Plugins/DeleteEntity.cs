// <copyright file="DeleteEntity.cs" company="">
// Copyright (c) 2014 All Rights Reserved
// </copyright>
// <author></author>
// <date>3/27/2014 8:57:41 AM</date>
// <summary>Implements the DeleteEntity Workflow Activity.</summary>
namespace Cloudrocket.Crm.Plugins
{
    using Cloudrocket.Xrm;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Workflow;
    using System;
    using System.Activities;
    using System.ServiceModel;

    public sealed class DeleteEntity : CodeActivity
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

            tracingService.Trace("Entered DeleteEntity.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("DeleteEntity.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // Courtesy of Gonzalo Ruiz https://crm2011workflowutils.codeplex.com/
                if (PrimaryEntity.Get(executionContext))
                {
                    t.Trace("Deleting process primary entity");
                    service.Delete(context.PrimaryEntityName, context.PrimaryEntityId);
                }
                else
                {
                    string relatedAttribute = RelatedAttributeName.Get(executionContext);
                    if (string.IsNullOrEmpty(relatedAttribute))
                    {
                        Helpers.Throw("If deleting related entity, related attribute name must be specified in the process delete step configuration");
                    }

                    // Retrieve primary entity with the required attribute
                    t.Trace("Retrieving process primary entity");
                    Entity primaryEntity = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(relatedAttribute));

                    if (primaryEntity.Contains(relatedAttribute))
                    {
                        EntityReference reference = primaryEntity[relatedAttribute] as EntityReference;
                        if (reference == null)
                        {
                            Helpers.Throw(string.Format("The attribute {0} on entity {1} is expected to be of EntityReference type",
                                relatedAttribute, context.PrimaryEntityName));
                        }

                        t.Trace("Deleting entity related to primary entity by attribute " + relatedAttribute);
                        service.Delete(reference.LogicalName, reference.Id);
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting DeleteEntity.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        #region Input Parameters

        [RequiredArgument]
        [Input("Delete the primary entity (True) or related entity (False)")]
        public InArgument<bool> PrimaryEntity { get; set; }

        [Input("If you're deleting the related entity, specify the attribute name that points to the related entity")]
        public InArgument<string> RelatedAttributeName { get; set; }

        #endregion Input Parameters

        #region Output Parameters

        #endregion Output Parameters
    }
}