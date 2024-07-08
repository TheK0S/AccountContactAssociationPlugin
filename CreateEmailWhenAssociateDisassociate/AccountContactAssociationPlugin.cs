using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.ServiceModel;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CreateEmailWhenAssociateDisassociate
{
    public class AccountContactAssociationPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is EntityReference)
                {
                    EntityReference relationshipReference = (EntityReference)context.InputParameters["Target"];
                    Relationship relationship = (Relationship)context.InputParameters["Relationship"];

                    tracingService.Trace($"relationship.SchemaName: {relationship.SchemaName}");

                    if (relationship.SchemaName != "kos_AccountsForContact") return;

                    EntityReferenceCollection relatedEntities = (EntityReferenceCollection)context.InputParameters["RelatedEntities"];
                    EntityReference accountReference = relatedEntities[0];
                    EntityReference contactReference = relationshipReference;

                    if (context.MessageName == "Associate")
                    {
                        CreateEmail(service, context.UserId, accountReference, contactReference, "Account has been associated", "Account has been added -");
                    }
                    else if (context.MessageName == "Disassociate")
                    {
                        CreateEmail(service, context.UserId, accountReference, contactReference, "Account has been disassociated", "Account has been removed -");
                    }
                }
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingService.Trace($"An error occurred: {ex.Message}");
                throw new InvalidPluginExecutionException($"An error occurred: {ex.Message}");
            }
            catch (Exception ex)
            {
                tracingService.Trace($"AccountContactAssociationPlugin: {ex.ToString()}");
                throw;
            }
        }

        private void CreateEmail(IOrganizationService service, Guid userId, EntityReference accountReference, EntityReference contactReference, string subjectTemplate, string descriptionTemplate)
        {
            Entity account = service.Retrieve("account", accountReference.Id, new ColumnSet("name"));
            Entity contact = service.Retrieve("contact", contactReference.Id, new ColumnSet("fullname"));

            string subject = $"{subjectTemplate} {account["name"]} with contact {contact["fullname"]}";
            string description = $"{descriptionTemplate} https://org82a3f762.crm11.dynamics.com/main.aspx?etn=account&id={accountReference.Id}&pagetype=entityrecord";

            Entity email = new Entity("email")
            {
                ["to"] = new EntityCollection(new[] { new Entity("activityparty") { ["partyid"] = contactReference } }),
                ["from"] = new EntityCollection(new[] { new Entity("activityparty") { ["partyid"] = new EntityReference("systemuser", userId) } }),
                ["subject"] = subject,
                ["description"] = description,
                ["regardingobjectid"] = contactReference
            };

            service.Create(email);
        }
    }
}
