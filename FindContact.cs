using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;

namespace WorkFlowTask
{
    public class FindContact : CodeActivity
    {
        [RequiredArgument]
        [Input("String input")]
        public InArgument<string> FieldName1StrIn { get; set; }

        [RequiredArgument]
        [Input("String input")]
        public InArgument<string> FieldName2StrIn { get; set; }

        [RequiredArgument]
        [Input("String input")]
        public InArgument<string> Param1StrIn { get; set; }

        [RequiredArgument]
        [Input("String input")]
        public InArgument<string> Param2StrIn { get; set; }

        [Output("Integer output")]
        public OutArgument<int> StatusIntOut { get; set; }

        [Output("EntityReference output")]
        [ReferenceTarget("contact")]
        public InOutArgument<EntityReference> ContactRefOut { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string param1 = Param1StrIn.Get(context);
            string param2 = Param2StrIn.Get(context);
            string field1 = FieldName1StrIn.Get(context);
            string field2 = FieldName2StrIn.Get(context);

            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);

            List<Entity> result = GetContactsByTwoParams(service, field1, param1, field2,param2);

            switch (result.Count) {
                case 0:
                    StatusIntOut.Set(context, 2);
                    ContactRefOut.Set(context, null);
                    break;
                case 1:
                    StatusIntOut.Set(context, 1);
                    ContactRefOut.Set(context, result[0]);
                    break;
                default:
                    StatusIntOut.Set(context, 3);
                    ContactRefOut.Set(context, result);
                    break;
            }
        }
        private List<Entity> GetContactsByTwoParams(IOrganizationService service, string field1, string param1, string field2,  string param2)
        {
            var result = new List<Entity>();
            var query = new QueryExpression("contact");

            query.Criteria.AddCondition(new ConditionExpression(field1, ConditionOperator.Equal, param1));
            query.Criteria.AddCondition(new ConditionExpression(field2, ConditionOperator.Equal, param2));
            query.Criteria.FilterOperator = LogicalOperator.And;

            result.AddRange(service.RetrieveMultiple(query).Entities.ToList());
            return result.ToList();
        }
    }
}
