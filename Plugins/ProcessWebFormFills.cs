// <copyright file="WebFormFill.cs" company="">Copyright (c) 2014 All Rights Reserved</copyright>
// <author></author>
// <date>2/25/2014 10:40:52 AM</date>
// <summary>
// Implements the WebFormFill Workflow Activity.
// </summary>
namespace Cloudrocket.Crm.Plugins
{
    using Cloudrocket.Xrm;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Workflow;
    using System;
    using System.Activities;
    using System.Linq;
    using System.ServiceModel;

    public sealed class ProcessWebFormFills : CodeActivity
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

            tracingService.Trace("Entered WebFormFill.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("WebFormFill.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            ITracingService t = tracingService;

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // TODO: Implement your custom Workflow business logic.

                #region 1. Get the Web Form Fill from the workflow context.

                t.Trace("1. Get the Form Fill from the workflow context.");
                cldrkt_webformfill webFormFill = (cldrkt_webformfill)service.Retrieve(
                    context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet { AllColumns = true });

                #endregion 1. Get the Web Form Fill from the workflow context.

                #region 2. Get the transaction owner and response email sender.

                QueryExpression userQuery = new QueryExpression {
                    EntityName = SystemUser.EntityLogicalName,
                    ColumnSet = new ColumnSet { AllColumns = true },
                    Criteria = new FilterExpression {
                        Conditions = {
                            new ConditionExpression {
                                AttributeName = "domainname",
                                Operator = ConditionOperator.Equal,
                                Values = { "Todd.Shelton@cloudrocket.com" },
                            }
                        }
                    }
                };
                t.Trace("2.1 Get the system user who will send the email.");
                SystemUser user = (SystemUser)service.RetrieveMultiple(userQuery).Entities.FirstOrDefault();
                t.Trace("2.2 The sender is: " + user.FullName.ToString());

                #endregion 2. Get the transaction owner and response email sender.

                #region 3. Look up the Contact from the email address, and create a new Contact if it doesn't already exist.

                t.Trace("3. Find or create the Contact from the email address." + webFormFill.cldrkt_Email);

                Contact contact = new Contact {
                    EMailAddress1 = webFormFill.cldrkt_Email,
                    FirstName = webFormFill.cldrkt_FirstName,
                    Id = Guid.NewGuid(),
                    LastName = webFormFill.cldrkt_LastName,
                    Telephone1 = webFormFill.cldrkt_BusinessPhone,
                };

                t.Trace("3.1 Look up the Contact using the email address entered: " + webFormFill.cldrkt_Email.ToString());

                QueryExpression contactsQuery = new QueryExpression {
                    EntityName = Contact.EntityLogicalName,
                    ColumnSet = new ColumnSet { AllColumns = true },
                    Criteria = new FilterExpression {
                        Conditions = {
                            new ConditionExpression {
                                AttributeName = "emailaddress1",
                                Operator = ConditionOperator.Equal,
                                Values = { contact.EMailAddress1 },
                            }
                        }
                    }
                };

                Contact c = (Contact)service.RetrieveMultiple(contactsQuery).Entities.FirstOrDefault();

                if (c != null)
                {
                    contact.Id = c.Id; // Will overwrite existing Contact data with entered data.
                    contact.ParentCustomerId = c.ParentCustomerId; // So it will be there for the Account lookup.
                    t.Trace("3.2.1 The existing contact is: " + contact.Id.ToString() + " " + contact.EMailAddress1);
                }
                else
                {
                    t.Trace("3.3.1 Create a new contact.");
                    contact.Id = service.Create(contact);
                    t.Trace("3.3.2 The new contact is: " + contact.Id.ToString() + " " + contact.EMailAddress1);
                }
                service.Update(contact);

                #endregion 3. Look up the Contact from the email address, and create a new Contact if it doesn't already exist.

                #region 4. Look up or create the Account and map this Contact to it.

                t.Trace("4. Look up or create the Account and map this Contact to it.");
                //t.Trace("4. Contact is " + contact.FullName);
                //t.Trace("4. Contact.Id is " + contact.Id);
                //t.Trace("4. contact.ParentCustomerId is " + contact.ParentCustomerId.ToString());

                Account account = new Account {
                    Name = webFormFill.cldrkt_Organization,
                };

                // Look up or create the parent Account.
                if (contact.ParentCustomerId != null)
                {
                    t.Trace("4.1 Build the parent account query.");

                    // Look up the  parent account.
                    QueryExpression parentAccountQuery = new QueryExpression {
                        EntityName = Account.EntityLogicalName,
                        ColumnSet = new ColumnSet { AllColumns = true },
                        Criteria = new FilterExpression {
                            Conditions = {
                                new ConditionExpression {
                                    AttributeName = "accountid",
                                    Operator = ConditionOperator.Equal,
                                    Values = { contact.ParentCustomerId.Id,},
                                }
                        },
                        },
                    };
                    t.Trace("4.2 Look up Account a.");

                    Account a = (Account)service.RetrieveMultiple(parentAccountQuery).Entities.FirstOrDefault();

                    t.Trace("4.3 If a exists, use it. Otherwise create a new Account.");

                    if (a != null)
                    {
                        t.Trace("4.3.1 The Account exists.");
                        account = a;
                        t.Trace("4.2.2 The existing Account is " + account.Name);
                    }
                    else
                    {
                        t.Trace("4.3.2 Create a new Account.");
                        account.Id = a.Id;
                        t.Trace("4.3.1 The new Account is " + account.Id.ToString());
                    }
                }
                else
                {
                    t.Trace("4.4 Create a new Account.");
                    account.Id = service.Create(account);
                };

                // Map the contact to the account.
                account.PrimaryContactId = new EntityReference {
                    Id = contact.Id,
                    LogicalName = Contact.EntityLogicalName,
                };
                service.Update(account);

                // Map the account to the contact.
                contact.ParentCustomerId = new EntityReference {
                    Id = account.Id,
                    LogicalName = Account.EntityLogicalName,
                };
                service.Update(contact);

                #endregion 4. Look up or create the Account and map this Contact to it.

                #region 5. Get the Campaign from the Campaign Activity ID and log a Campaign Response.

                t.Trace("5. Get the Campaign Activity, if any...");
                CampaignActivity campaignActivity = new CampaignActivity();
                CampaignResponse campaignResponse = new CampaignResponse();

                Guid campaignActivityId = Guid.Empty;

                t.Trace("5.1 webFormFill.cldrkt_CampaignActivityID: " + webFormFill.cldrkt_CampaignActivityID);
                if (String.IsNullOrWhiteSpace(webFormFill.cldrkt_CampaignActivityID))
                {
                    campaignActivityId = Guid.Empty;
                }
                else
                {
                    t.Trace("5.2 We have a webFormFill.cldrkt_CampaignActivityID: " + webFormFill.cldrkt_CampaignActivityID);

                    Guid.TryParse(webFormFill.cldrkt_CampaignActivityID, out campaignActivityId);

                    t.Trace("5.2.1 CampaignActivityID is " + campaignActivityId.ToString());

                    if (campaignActivityId != Guid.Empty)
                    {
                        t.Trace("5.2.2 Look up the Campaign Activity...");
                        campaignActivity = (CampaignActivity)service.Retrieve(
                            CampaignActivity.EntityLogicalName, campaignActivityId, new ColumnSet { AllColumns = true });

                        t.Trace("5.2.3 campaignActivityId: " + campaignActivityId);
                        t.Trace("5.2.4 campaignActivity.Id: " + campaignActivity.Id.ToString());

                        if (campaignActivity != null) // Found a Campaign Activity.
                        {
                            // Create a Campaign Response.
                            t.Trace("5.3 Create a Campaign Response...");

                            campaignResponse.ChannelTypeCode = new OptionSetValue((int)636280001); // 636280001: Web Page Form fill

                            campaignResponse.Customer = new ActivityParty[] {
                                new ActivityParty { PartyId = new EntityReference(Contact.EntityLogicalName, contact.Id) }
                            };

                            campaignResponse.FirstName = webFormFill.cldrkt_FirstName;
                            campaignResponse.LastName = webFormFill.cldrkt_LastName;
                            campaignResponse.EMailAddress = webFormFill.cldrkt_Email;
                            campaignResponse.Telephone = webFormFill.cldrkt_BusinessPhone;
                            campaignResponse.CompanyName = webFormFill.cldrkt_Organization;
                            campaignResponse.PromotionCodeName = webFormFill.cldrkt_PromotionCode;

                            campaignResponse.cldrkt_CampaignActivityId = new EntityReference {
                                Id = campaignActivity.Id,
                                LogicalName = CampaignActivity.EntityLogicalName,
                            };
                            campaignResponse.OriginatingActivityId = new EntityReference {
                                Id = webFormFill.Id,
                                LogicalName = cldrkt_webformfill.EntityLogicalName,
                            };
                            campaignResponse.RegardingObjectId = new EntityReference // Required, must be the parent campaign
                            {
                                Id = campaignActivity.RegardingObjectId.Id,
                                LogicalName = Campaign.EntityLogicalName,
                            };
                            
                            campaignResponse.ReceivedOn = webFormFill.CreatedOn;

                            campaignResponse.Subject = webFormFill.Subject; //TODO: Change to an available field.

                            t.Trace("5.2.5 Create the Campaign Response.");

                            campaignResponse.ActivityId = service.Create(campaignResponse);
                            t.Trace("5.3.1 campaignResponse.ActivityId: " + campaignResponse.ActivityId);
                            t.Trace("5.3.2 campaignResponse.Id: " + campaignResponse.Id.ToString());

                            // Update the Campaign Response.
                            t.Trace("5.4 Update the Campaign Response.");

                            if (campaignResponse.Id != Guid.Empty)
                            {
                                service.Update(campaignResponse);
                                t.Trace("5.4.1 campaignResponse.Id = " + campaignResponse.Id.ToString());
                            }

                            // Add the Campaign Activity to the Web Form Fill.
                            t.Trace("5.5. Add the Campaign Activity to the Web Form fill");

                            webFormFill.cldrkt_Campaign = new EntityReference {
                                Id = campaignActivity.RegardingObjectId.Id,
                                LogicalName = campaignActivity.RegardingObjectId.LogicalName,
                            };
                            webFormFill.cldrkt_CampaignActivity = new EntityReference {
                                Id = campaignActivity.Id,
                                LogicalName = campaignActivity.LogicalName,
                            };
                            webFormFill.cldrkt_CampaignResponse = new EntityReference {
                                Id = campaignResponse.Id,
                                LogicalName = campaignResponse.LogicalName,
                            };
                            t.Trace("5.6 Update the webFormFill.");
                            service.Update(webFormFill);
                        }
                    }
                }

                #endregion 5. Get the Campaign from the Campaign Activity ID and log a Campaign Response.

                #region 6. Create a new Opportunity and map it to the Contact.

                t.Trace("6. Create a new Opportunity and map it to the Contact. ");

                string productNumber =  // Defaulting to SMSP.  The Product Number has to be valid.
                    String.IsNullOrEmpty(webFormFill.cldrkt_ProductNumber) ? "SMSP-License" : webFormFill.cldrkt_ProductNumber;

                QueryExpression productQuery = new QueryExpression {
                    EntityName = Product.EntityLogicalName,
                    ColumnSet = new ColumnSet { AllColumns = true },
                    Criteria = new FilterExpression {
                        Conditions = {
                              new ConditionExpression {
                                  AttributeName = "productnumber",
                                  Operator = ConditionOperator.Equal,
                                  Values = { productNumber },
                              }
                          }
                    }
                };

                t.Trace("6.1.1 Look up the product. ");

                Product product = (Product)service.RetrieveMultiple(productQuery).Entities.FirstOrDefault();

                t.Trace("6.1.2 product.Id is " + product.Id.ToString() + " product.ProductId is " + product.ProductId);

                t.Trace("6.1.3 product.ProductId is " + product.Id.ToString() + " ");

                t.Trace("6.2 Create the Opportunity. ");
                t.Trace("6.2.0 campaignActivity.Id is " + campaignActivity.Id.ToString());
                t.Trace("6.2.1 campaignActivity.RegardingObjectId.Id is " + campaignActivity.RegardingObjectId.Id.ToString());
                t.Trace("6.2.2 account.Name and product.ProductNumber are " + account.Name + " " + product.ProductNumber);
                t.Trace("6.2.3  product.PriceLevelId is " + product.PriceLevelId.Id.ToString());

                Opportunity opportunity = new Opportunity {
                    CampaignId = campaignActivity.RegardingObjectId,
                    cldrkt_EstimatedUsers = (int?)webFormFill.cldrkt_ProductQuantity,
                    Name = webFormFill.Subject, // Required. 
                    cldrkt_DateofLastContact = webFormFill.CreatedOn,
                    IsRevenueSystemCalculated = true,
                    ParentAccountId = new EntityReference {
                        Id = account.Id,
                        LogicalName = Account.EntityLogicalName,
                    },
                    ParentContactId = new EntityReference {
                        Id = contact.Id,
                        LogicalName = Contact.EntityLogicalName,
                    },
                    PriceLevelId = product.PriceLevelId, // Required
                    StepName = "1-Conversation",
                    TransactionCurrencyId = product.TransactionCurrencyId, // Required.
                };

                t.Trace("6.2.5 opportunity.TransactionCurrencyId is " + opportunity.TransactionCurrencyId.Name.ToString());
                t.Trace("6.2.6 TransactionCurrencyId.Id is " + opportunity.TransactionCurrencyId.Id.ToString());
                t.Trace("6.2.6.1 opportunity.ParentContactId.Id is " + opportunity.ParentContactId.Id.ToString());

                opportunity.Id = service.Create(opportunity);
                service.Update(opportunity);

                t.Trace("6.2.7 opportunity.Id is " + opportunity.Id.ToString());
                t.Trace("6.2.7.1 ShowMe price is " + Helpers.GetShowMePricePerUser((decimal)webFormFill.cldrkt_ProductQuantity));

                t.Trace("6.3 Create the OpportunityProduct.");
                OpportunityProduct opportunityProduct = new OpportunityProduct {
                    OpportunityId = new EntityReference {
                        LogicalName = Opportunity.EntityLogicalName,
                        Id = opportunity.Id,
                    },
                    ProductId = new EntityReference {
                        LogicalName = Product.EntityLogicalName,
                        Id = product.Id,
                    },
                    UoMId = new EntityReference {
                        LogicalName = UoM.EntityLogicalName,
                        Id = product.DefaultUoMId.Id,
                    },
                    Quantity = webFormFill.cldrkt_ProductQuantity,
                    PricePerUnit = new Money {
                        Value = Helpers.GetShowMePricePerUser((decimal)webFormFill.cldrkt_ProductQuantity),
                    },
                    IsPriceOverridden = true,
                };

                t.Trace("6.3.1 Creating the opportunityProduct. ");
                opportunityProduct.Id = service.Create(opportunityProduct);

                t.Trace("6.3.2 opportunityProduct.Id is " + opportunityProduct.Id.ToString());
                t.Trace("6.3.3 opportunityProductProductId.Id is " + opportunityProduct.ProductId.Id.ToString());

                t.Trace("6.3.4 opportunityProduct.Quantity is " + opportunityProduct.Quantity);
                t.Trace("6.3.5 opportunityProduct.Quantity.Value is " + opportunityProduct.Quantity.Value);
                t.Trace("6.3.6 opportunityProduct.PricePerUnit is " + opportunityProduct.PricePerUnit);
                t.Trace("6.3.7 opportunityProduct.PricePerUnit.Value is " + opportunityProduct.PricePerUnit.Value);

                service.Update(opportunityProduct);
                service.Update(opportunity);

                #endregion 6. Create a new Opportunity and map it to the Contact.

                #region 7. Get the response email template.

                t.Trace(" 7. Get the email template from the Web Form Fill, otherwise use a default template");
                QueryExpression templateQuery = new QueryExpression {
                    EntityName = Template.EntityLogicalName,
                    ColumnSet = new ColumnSet { AllColumns = true },
                    Criteria = new FilterExpression {
                        Conditions = {
                            new ConditionExpression {
                                AttributeName = "title",
                                Operator = ConditionOperator.Equal,
                                Values = { webFormFill.cldrkt_EmailTemplateTitle },
                            }
                        }
                    }
                };

                Template emailTemplate = new Template();
                Guid defaultEmailTemplateId = Guid.Parse("d4fe12fd-72d2-e311-9e62-6c3be5be5e68"); // Default, SMSP demo request
                Guid emailTemplateId = new Guid();

                if (String.IsNullOrEmpty(webFormFill.cldrkt_EmailTemplateTitle))
                {
                    emailTemplateId = defaultEmailTemplateId;
                    t.Trace("7.1 No email template set from the web form.");
                }
                else
                {
                    t.Trace("7.2.1 Looking up Template from webFormFill: " + webFormFill.cldrkt_EmailTemplateTitle);

                    emailTemplate = (Template)service.RetrieveMultiple(templateQuery).Entities.FirstOrDefault();
                    if (emailTemplate == null)
                    { t.Trace("Template is null"); }
                    else
                    {
                        t.Trace("Template is not null.");
                        t.Trace("Template type is: " + emailTemplate.TemplateTypeCode.ToString());
                    }

                    t.Trace("7.2.1 Looked up Template using the Title. ");

                    emailTemplateId = emailTemplate == null ? defaultEmailTemplateId : emailTemplate.Id;
                    t.Trace("7.2.2 emailTemplateId: " + emailTemplateId.ToString());
                }

                t.Trace("7.3.1 The email template is " + emailTemplate.Title.ToString() + " type of " + emailTemplate.TemplateTypeCode + " Id: " + emailTemplateId.ToString());

                #endregion 7. Get the response email template.

                #region 8. Create and send the response email.

                t.Trace("8. Create and send the email message.");
                t.Trace("8. Send from: " + user.FullName.ToString());
                t.Trace("8. Send to: " + contact.Id.ToString() + " using template " + emailTemplate.Title + " with Id " + emailTemplateId.ToString());
                // Create an email using an Opportunity template. "To" is a Contact type.
                SendEmailFromTemplateRequest emailUsingTemplateReq = new SendEmailFromTemplateRequest {
                    Target = new Email {
                        To = new ActivityParty[] { new ActivityParty { PartyId = new EntityReference(Contact.EntityLogicalName, opportunity.ParentContactId.Id) } },
                        From = new ActivityParty[] { new ActivityParty { PartyId = new EntityReference(SystemUser.EntityLogicalName, user.Id) } },
                        Subject = "",
                        Description = "",
                        DirectionCode = true,
                    },
                    RegardingId = opportunity.Id, // Required, and the type must match the Email Template type.
                    RegardingType = emailTemplate.TemplateTypeCode,

                    TemplateId = emailTemplateId,
                };

                t.Trace("8.1 Send email to: " + opportunity.ParentContactId.Id.ToString() + " from: " + user.DomainName);
                t.Trace("8.1.1 Contact ID is: " + contact.Id.ToString() + ", email template is " + emailTemplate.Id.ToString() + ", opportunity is " + opportunity.Id.ToString());
                t.Trace("8.1.2 email template id is: " + emailUsingTemplateReq.TemplateId.ToString() );

                SendEmailFromTemplateResponse email = (SendEmailFromTemplateResponse)service.Execute(emailUsingTemplateReq);

                t.Trace("8.2 Email sent: " + email.Id.ToString());

                #endregion 8. Create and send the response email.

                #region 9. Add this Contact to the Marketing List, and create the list if it doesn't exist.

                t.Trace("9. Add this Contact to the Marketing List. " + contact.Id.ToString() + " to List " + webFormFill.cldrkt_AddToMarketingList);

                List staticContactList = new List {
                    CreatedFromCode = new OptionSetValue((int)2), // Required.  Account = 1, Contact = 2, Lead = 4.
                    Id = Guid.NewGuid(), // Required.
                    ListName = webFormFill.cldrkt_AddToMarketingList, // Required.
                    LogicalName = List.EntityLogicalName,
                    OwnerId = new EntityReference { // Required.
                        Id = user.Id,
                        LogicalName = SystemUser.EntityLogicalName,
                    },
                    StatusCode = new OptionSetValue((int)0),
                    Type = false, // Required.  True = dynamic, false = static.
                };

                QueryExpression listQuery = new QueryExpression {
                    EntityName = List.EntityLogicalName,
                    ColumnSet = new ColumnSet { AllColumns = true },
                    Criteria = new FilterExpression {
                        Conditions = {
                            new ConditionExpression {
                                AttributeName = "listname",
                                Operator = ConditionOperator.Equal,
                                Values = { webFormFill.cldrkt_AddToMarketingList},
                            }
                        }
                    }
                };
                t.Trace("9.1 Get this list, if it exists: " + webFormFill.cldrkt_AddToMarketingList);

                Entity list = service.RetrieveMultiple(listQuery).Entities.FirstOrDefault();
                t.Trace("9.2 Look up the list.");

                if (list == null)
                {
                    t.Trace("9.3.1 Create a new list: " + staticContactList.Id.ToString());
                    staticContactList.Id = service.Create(staticContactList);
                }
                else
                {
                    t.Trace("9.3.2 Use the list we found: " + list.Id.ToString());
                    staticContactList.Id = list.Id;
                }

                t.Trace("9.4 Add the Contact " + contact.Id.ToString() + " to List " + staticContactList.Id.ToString());
                AddMemberListRequest addMemberListRequest = new AddMemberListRequest {
                    EntityId = contact.Id,
                    ListId = staticContactList.Id,
                };

                service.Execute(addMemberListRequest);

                #endregion 9. Add this Contact to the Marketing List, and create the list if it doesn't exist.

                #region 10. Update the entities we've worked on.

                t.Trace("10. Update the entities we've worked on. ");

                webFormFill.RegardingObjectId = new EntityReference { Id = contact.Id, LogicalName = Contact.EntityLogicalName, };
                service.Update(webFormFill);

                service.Update(contact);
                service.Update(opportunityProduct);
                service.Update(opportunity);
                service.Update(webFormFill);

                #endregion 10. Update the entities we've worked on.

                //throw new InvalidPluginExecutionException("Finished processing the Web Form Fill update.");
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting WebFormFill.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        // Define Input/Output Arguments [RequiredArgument] [Input("Web Form Fill")]
        // [ReferenceTarget("cldrkt_webformfill")] public InArgument<EntityReference>
        // inputWebFormFill { get; set; }

        // Use this section to set up output parameters to set on workflow form.
        //[Output("OptionSet Day Of The Week output")]
        //[AttributeTarget("cldrkt_pageview", "cldrkt_createdonweekdayoptionset")]
        // public OutArgument<OptionSetValue> outputOptionSetValueDayOfTheWeek { get; set; }
    }
}