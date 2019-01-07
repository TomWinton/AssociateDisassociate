using System.Activities;

// These namespaces are found in the Microsoft.Xrm.Sdk.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

// These namespaces are found in the Microsoft.Xrm.Sdk.Workflow.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Xrm.Sdk.Workflow;

// These namespaces are found in the Microsoft.Crm.Sdk.Proxy.dll assembly
// located in the SDK\bin folder of the SDK download.
using Microsoft.Xrm.Sdk.Messages;

namespace AssociateDisassociate
{

    public sealed partial class AssociateDisassociate : CodeActivity
    {

        //The primary entity input details
        [Input("Contact")]
        [ReferenceTarget("contact")]
        public InArgument<EntityReference> inputContact
        {
            get;
            set;
        }

        //The related entity input details
        [Input("Account")]
        [ReferenceTarget("account")]
        public InArgument<EntityReference> inputAccount
        {
            get;
            set;
        }

        //The boolean choice to determine whether to associate or disassociate these records
        [Input("T=disassociate F=associate")]
        [Output("result")]
        [Default("True")]
        public InOutArgument<bool> Bool
        {
            get;
            set;
        }

        protected override void Execute(CodeActivityContext executionContext)
        {

            //Build the connection
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            string contactId = this.inputContact.Get(executionContext).Id.ToString(); // Guid of primary entity to be used in fetch
            string accountId = this.inputAccount.Get(executionContext).Id.ToString(); // Guid of secondary entity to be used in fetch

            {
                //The fetch to be used if disassociating
                string fetchxml = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                                        <entity name='contact'>
                                                         <attribute name='contactid' />
                                                          <order attribute='contactid' descending='false' />
                                                           <filter type='and'>
                                                                <condition attribute='contactid' operator='eq' uitype='contact' value='{" + contactId + @"}' />
                                                           </filter>
                                                            <link-entity name='gg_contact_account_entitypermission' from='contactid' to='contactid' visible='false' intersect='true'>
                                                              <link-entity name='account' from='accountid' to='accountid' alias='ac'>
                                                                <filter type='and'>
                                                                 <condition attribute='accountid' operator='eq' uitype='account' value='{" + accountId + @"}' />
                                                            </filter>
                                                                 </link-entity>
                                                             </link-entity>
                                                             </entity>
                                                        </fetch>";


                bool TrueFalse = this.Bool.Get(executionContext); //result of the yes no input determing whether to associate or disaasociate

                
                    EntityCollection collRecords = service.RetrieveMultiple(new FetchExpression(fetchxml));
                if (TrueFalse) //Logic for disassociate
                {
                    if (collRecords != null && collRecords.Entities != null && collRecords.Entities.Count > 0)
                    {
                        EntityReferenceCollection collection = new EntityReferenceCollection();
                        foreach (var entity in collRecords.Entities)
                        {
                            var reference = new EntityReference("list", entity.Id);
                            collection.Add(reference); //Create a collection of entity references
                        }
                        Relationship relationship = new Relationship("gg_contact_account_entitypermission"); //schema name of N:N relationship
                        service.Disassociate("account", this.inputAccount.Get(executionContext).Id, relationship, collection); //Pass the entity reference collections to be disassociated from the specific Email Send record
                    }
                }
                else //Logic for associate
                {
                    if (collRecords.Entities.Count == 0) { 
                        AssociateRequest request = new AssociateRequest
                    {

                        Target = new EntityReference("account", this.inputAccount.Get(executionContext).Id),
                        RelatedEntities = new EntityReferenceCollection {
                            new EntityReference("list", this.inputContact.Get(executionContext).Id)
                        },
                        Relationship = new Relationship("gg_contact_account_entitypermission")
                    };

                    service.Execute(request);
                    }
                }

            }

        }

    }

}


