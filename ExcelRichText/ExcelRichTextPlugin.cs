using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Text;

namespace PluginSandbox
{
    public class ExcelRichTextPlugin : IPlugin
    {
        const string MSG_EXPORT_TO_EXCEL = "ExportToExcel";
        const string PARAM_BUSINESS_ENTITY_COLLECTION = "BusinessEntityCollection";
        
        /// <summary>
        /// Called when the plugin step is triggered.
        /// </summary>
        /// <param name="serviceProvider">IServiceProvider passed by the execution pipeline.</param>
        public void Execute(IServiceProvider serviceProvider)
        {
            // Retrieving the context and OrganizationService
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService orgService = serviceFactory.CreateOrganizationService(context.UserId);

            // Checking if the plugin is running in the expected context
            if (!IsExportExcel(context) || !context.OutputParameters.ContainsKey(PARAM_BUSINESS_ENTITY_COLLECTION))
                return;

            // Returning immediately if the current query returns no entity
            var entities = context.OutputParameters[PARAM_BUSINESS_ENTITY_COLLECTION] as EntityCollection;
            if (entities.Entities.Count < 1)
                return;

            // Retrieving the attributes containing HTML
            var metadata = orgService.GetEntityMetadata(context.PrimaryEntityName);
            var richTextAttributes = GetRichTextAttributes(metadata);
            if (richTextAttributes.Count < 1)
                return;

            // Removing the HTML in all the Rich Text attributes
            foreach (var entity in entities.Entities)
            {
                foreach (var att in richTextAttributes)
                {
                    string richText;
                    if (entity.TryGetAttributeValue<string>(att, out richText))
                    {
                        entity[att] = HTMLToText(richText);
                    }
                }
            }
        }

        /// <summary>
        /// Detects if the plugin has been triggered by an Export to Excel action.
        /// </summary>
        /// <param name="context">The current plugin context.</param>
        /// <returns>true if an Export to Excel has triggered this plugin; otherwise false.</returns>
        private bool IsExportExcel(IPluginExecutionContext context)
        {
            if (context.MessageName == MSG_EXPORT_TO_EXCEL)
                return true;
            else if (context.ParentContext != null)
                return IsExportExcel(context.ParentContext);
            else
                return false;
        }

        /// <summary>
        /// Returns the list of the attributes that may contain HTML for the specified entity.
        /// </summary>
        /// <param name="metadata">The entity's metadata.</param>
        /// <returns>An hashset of the attributes logical names.</returns>
        private HashSet<string> GetRichTextAttributes(EntityMetadata metadata)
        {
            var richTextAttributes = new HashSet<string>();
            foreach (var attribute in metadata.Attributes)
            {
                if ((attribute is StringAttributeMetadata && ((StringAttributeMetadata)attribute).FormatName == "RichText") ||
                    (attribute is MemoAttributeMetadata && (((MemoAttributeMetadata)attribute).FormatName == "RichText")))
                {
                    richTextAttributes.Add(attribute.LogicalName);
                }
            }
            return richTextAttributes;
        }

        /// <summary>
        /// Converts HTML to plain text by removing the HTML tags and formatting.
        /// </summary>
        /// <param name="HTMLCode">The text to convert.</param>
        /// <returns>The specified content as plain text.</returns>
        private string HTMLToText(string HTMLCode)
        {
            // Source : https://beansoftware.com/ASP.NET-Tutorials/Convert-HTML-To-Plain-Text.aspx
            // Remove new lines since they are not visible in HTML
            HTMLCode = HTMLCode.Replace("\n", " ");

            // Remove tab spaces
            HTMLCode = HTMLCode.Replace("\t", " ");

            // Remove multiple white spaces from HTML
            HTMLCode = Regex.Replace(HTMLCode, "\\s+", " ");

            // Remove HEAD tag
            HTMLCode = Regex.Replace(HTMLCode, "<head.*?</head>", ""
                                , RegexOptions.IgnoreCase | RegexOptions.Singleline);

            // Remove any JavaScript
            HTMLCode = Regex.Replace(HTMLCode, "<script.*?</script>", ""
              , RegexOptions.IgnoreCase | RegexOptions.Singleline);

            // Replace special characters like &, <, >, " etc.
            StringBuilder sbHTML = new StringBuilder(HTMLCode);
            // Note: There are many more special characters, these are just
            // most common. You can add new characters in this arrays if needed
            string[] OldWords = {"&nbsp;", "&amp;", "&quot;", "&lt;",
   "&gt;", "&reg;", "&copy;", "&bull;", "&trade;"};
            string[] NewWords = { " ", "&", "\"", "<", ">", "Â®", "Â©", "â€¢", "â„¢" };
            for (int i = 0; i < OldWords.Length; i++)
            {
                sbHTML.Replace(OldWords[i], NewWords[i]);
            }

            // Check if there are line breaks (<br>) or paragraph (<p>)
            sbHTML.Replace("<br>", "\n<br>");
            sbHTML.Replace("<br ", "\n<br ");
            sbHTML.Replace("<p ", "\n<p ");

            // Finally, remove all HTML tags and return plain text
            return System.Text.RegularExpressions.Regex.Replace(
              sbHTML.ToString(), "<[^>]*>", "");
        }
    }
}
