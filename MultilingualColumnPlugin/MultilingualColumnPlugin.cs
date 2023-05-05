using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WWW.Samples.Plugins
{
    public class MultilingualColumnPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService orgService = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.SharedVariables.ContainsKey("WWW.Samples.Plugin.MultilingualColumnPlugin"))
                return;
            context.SharedVariables.Add("WWW.Samples.Plugin.MultilingualColumnPlugin", true);

            int userLocale = GetUserLocale(orgService, context.UserId);

            if (string.Equals(context.MessageName, "retrieve", StringComparison.InvariantCultureIgnoreCase))
            {
                var target = context.OutputParameters["BusinessEntity"] as Entity;
                if (target != null)
                {
                    var muiMetadata = GetMuiMetadata(orgService, target.LogicalName);
                    TranslateMuiRecord(orgService, target, userLocale, muiMetadata, true);

                    foreach (var att in target.Attributes)
                    {
                        if (att.Value is EntityReference)
                        {
                            var etnRef = (EntityReference)att.Value;
                            TranslateMuiReference(orgService, etnRef, userLocale);
                        }
                    }
                }
            }
            else if (string.Equals(context.MessageName, "retrievemultiple", StringComparison.InvariantCultureIgnoreCase))
            {
                var target = context.OutputParameters["BusinessEntityCollection"] as EntityCollection;
                if (target != null)
                {
                    var muiMetadata = GetMuiMetadata(orgService, target.EntityName);
                    foreach (var record in target.Entities)
                    {
                        TranslateMuiRecord(orgService, record, userLocale, muiMetadata);
                    }

                    foreach (var record in target.Entities)
                    {
                        foreach (var att in record.Attributes)
                        {
                            if (att.Value is EntityReference)
                            {
                                var etnRef = (EntityReference)att.Value;
                                TranslateMuiReference(orgService, etnRef, userLocale);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Replaces the primary name attribute's value with the localized value.
        /// </summary>
        /// <param name="orgService">The service.</param>
        /// <param name="record">The record to localize.</param>
        /// <param name="localeId">The locale id.</param>
        /// <param name="muiMetadata">Contains the names of the localized attributes.</param>
        /// <param name="optimistic">When set to true the localized values will be retrieved from the record if available. When set to false a query will be executed to retrieve the localized attributes' values.</param>
        private void TranslateMuiRecord(IOrganizationService orgService, Entity record, int localeId, MuiMetadata muiMetadata, bool optimistic = false)
        {
            if (muiMetadata != null)
            {
                if(optimistic == true && record.Attributes.ContainsKey(muiMetadata.LocalizedNameAttributes[localeId]))
                {
                    record[muiMetadata.PrimaryNameAttribute] = record.GetAttributeValue<string>(muiMetadata.LocalizedNameAttributes[localeId]);
                }
                else
                {
                    var localizedRecord = orgService.Retrieve(record.LogicalName, record.Id, new ColumnSet(new[] { muiMetadata.LocalizedNameAttributes[1033], muiMetadata.LocalizedNameAttributes[1036] }));
                    record[muiMetadata.PrimaryNameAttribute] = localizedRecord.GetAttributeValue<string>(muiMetadata.LocalizedNameAttributes[localeId]);
                }
            }
        }

        private void TranslateMuiReference(IOrganizationService orgService, EntityReference entityReference, int localeId)
        {
            var muiMetadata = GetMuiMetadata(orgService, entityReference.LogicalName);
            if (muiMetadata != null)
            {
                var localizedRecord = orgService.Retrieve(entityReference.LogicalName, entityReference.Id, new ColumnSet(new[] { muiMetadata.LocalizedNameAttributes[1033], muiMetadata.LocalizedNameAttributes[1036] }));
                entityReference.Name = localizedRecord.GetAttributeValue<string>(muiMetadata.LocalizedNameAttributes[localeId]);
            }
        }

        private MuiMetadata GetMuiMetadata(IOrganizationService orgService, string logicalName)
        {
            var metadata = orgService.GetEntityMetadata(logicalName);
            var primaryName = metadata.PrimaryNameAttribute;

            var name1033 = metadata.Attributes.FirstOrDefault(a => a.LogicalName == $"{primaryName}_1033");
            var name1036 = metadata.Attributes.FirstOrDefault(a => a.LogicalName == $"{primaryName}_1036");

            if (name1033 != null && name1036 != null)
            {
                return new MuiMetadata
                {
                    PrimaryNameAttribute = primaryName,
                    LocalizedNameAttributes = new Dictionary<int, string>()
                    {
                        { 1033, name1033.LogicalName },
                        { 1036, name1036.LogicalName }
                    }
                };
            }
            return null;
        }

        private int GetUserLocale(IOrganizationService orgService, Guid userId)
        {
            var userSettings = orgService.Retrieve("usersettings", userId, new ColumnSet(new[] { "uilanguageid" }));
            return userSettings.GetAttributeValue<int>("uilanguageid");
        }

        private class MuiMetadata
        {
            public string PrimaryNameAttribute { get; set; }

            public Dictionary<int, string> LocalizedNameAttributes { get; set; }
        }
    }
}
