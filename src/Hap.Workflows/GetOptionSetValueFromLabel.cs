using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;

namespace Hap.Workflows
{
    public class GetOptionSetValueFromLabel : CodeActivity
    {


        [RequiredArgument]
        [Input("Text")]
        public InArgument<string> Text { get; set; }

        [RequiredArgument]
        [Input("EntityName")]
        public InArgument<string> EntityName { get; set; }

        [RequiredArgument]
        [Input("FieldName")]
        public InArgument<string> FieldName { get; set; }

        [Output("OptionSetValue")]
        public OutArgument<int> OptionSetValue { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            var context = executionContext.GetExtension<IWorkflowContext>();
            var serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            var organizationService = serviceFactory.CreateOrganizationService(null);

            var text = Text.Get(executionContext);
            var entityName = EntityName.Get(executionContext);
            var fieldName = FieldName.Get(executionContext);
            var optionSetValue = GetOptionSetValue(organizationService, entityName, fieldName, text);

            executionContext.SetValue(OptionSetValue, optionSetValue);
        }

        public int GetOptionSetValue(IOrganizationService service, string entityName, string fieldName, string text)
        {
            try
            {
                var retrieveAttributeRequest = new RetrieveAttributeRequest
                {
                    EntityLogicalName = entityName,
                    LogicalName = fieldName,
                    RetrieveAsIfPublished = true
                };

                var retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
                var retrievedPicklistAttributeMetadata = (PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
                OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
                foreach (OptionMetadata oMD in optionList)
                {
                    if (oMD.Label.LocalizedLabels[0].Label.ToString().ToLower() == text.ToLower())
                    {
                        var selectedOptionValue = oMD.Value.Value;
                        return selectedOptionValue;

                    }
                }
                return -1;
            }
            catch (System.Exception)
            {
                return -1;
            }

        }
    }
}
