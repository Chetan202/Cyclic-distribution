using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Cyclic_distribution
{
    public class Class1 : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Get the context, tracing service, and organization service from the service provider
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            tracingService.Trace("AssignCaseToUserUsingHashing Plugin execution started.");

            // Check if the Target parameter exists and is of type Entity
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity caseEntity = (Entity)context.InputParameters["Target"];
                tracingService.Trace("Target entity is: {0}", caseEntity.LogicalName);

                // Ensure the entity is a case (incident)
                if (caseEntity.LogicalName != "incident")
                {
                    tracingService.Trace("Entity is not a case. Exiting plugin.");
                    return;
                }

                try
                {
                    tracingService.Trace("Retrieving all active users with the specified criteria.");

                    // FetchXML to retrieve users based on provided criteria (e.g., positionid)
                    string fetchXml = @"
            <fetch version='1.0' output-format='xml-platform' mapping='logical' no-lock='false' distinct='true'>
                <entity name='systemuser'>
                    <attribute name='fullname'/>
                    <attribute name='systemuserid'/>
                    <filter type='and'>
                        <condition attribute='isdisabled' operator='eq' value='0'/>
                        <condition attribute='accessmode' operator='eq' value='0'/>
                        <condition attribute='positionid' operator='eq' value='{4f0b53fa-cf74-ef11-ac20-6045bdc5d905}' uiname='Associate' uitype='position'/>
                    </filter>
                </entity>
            </fetch>";

                    // Retrieve the users based on the FetchXML
                    EntityCollection users = service.RetrieveMultiple(new FetchExpression(fetchXml));
                    tracingService.Trace("Total active users retrieved: {0}", users.Entities.Count);

                    if (users.Entities.Count > 0)
                    {
                        // Get the case ID for hashing logic
                        Guid caseId = caseEntity.Id;
                        tracingService.Trace("Case ID: {0}", caseId);

                        // Hash the case ID and distribute the cases cyclically among users
                        int userIndex = GetUserIndexByHashing(caseId, users.Entities.Count);
                        tracingService.Trace("User index calculated from hash: {0}", userIndex);

                        Entity selectedUser = users.Entities[userIndex];
                        tracingService.Trace("Assigning case to user: {0}", selectedUser["fullname"]);

                        // Update the case with the selected user as the owner
                        Entity updatedCaseEntity = new Entity(caseEntity.LogicalName, caseEntity.Id);
                        updatedCaseEntity["ownerid"] = new EntityReference("systemuser", selectedUser.Id);
                        service.Update(updatedCaseEntity);

                        tracingService.Trace("Case assigned to user: {0}", selectedUser["fullname"]);
                    }
                    else
                    {
                        tracingService.Trace("No active users found based on the FetchXML criteria.");
                    }
                }
                catch (Exception ex)
                {
                    tracingService.Trace("Error in AssignCaseToUserUsingHashing: {0}", ex.ToString());
                    throw new InvalidPluginExecutionException("An error occurred in the AssignCaseToUserUsingHashing plugin.", ex);
                }
            }
            else
            {
                tracingService.Trace("Target is not an entity.");
            }

            tracingService.Trace("AssignCaseToUserUsingHashing Plugin execution ended.");
        }

        // Helper method to calculate the index of the user to assign the case to using hashing logic
        private int GetUserIndexByHashing(Guid caseId, int totalUsers)
        {
            // Get the hash code of the case ID
            int hash = caseId.GetHashCode();

            // Ensure the hash value is positive
            if (hash < 0) hash = -hash;

            // Calculate the user index by taking the modulus of the hash with the total number of users
            int userIndex = hash % totalUsers;

            return userIndex;
        }
    }
}