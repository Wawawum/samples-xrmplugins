using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Extensions;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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
                        entity[att] = RemoveHtml(richText);
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
        /// <param name="input">The text to convert.</param>
        /// <returns>The specified content as plain text.</returns>
        public static string RemoveHtml(string input)
        {
            string output = Regex.Replace(input, "<.*?>", string.Empty); // Remove HTML tags
            output = Regex.Replace(output, @"&nbsp;|&#160;", " "); // Replace &nbsp; and &#160; with a space
            output = Regex.Replace(output, @"&amp;|&#38;", "&"); // Replace &amp; and &#38; with &
            output = Regex.Replace(output, @"&lt;|&#60;", "<"); // Replace &lt; and &#60; with <
            output = Regex.Replace(output, @"&gt;|&#62;", ">"); // Replace &gt; and &#62; with >
            output = Regex.Replace(output, @"&quot;|&#34;", "\""); // Replace &quot; and &#34; with "
            output = Regex.Replace(output, @"&apos;|&#39;", "'"); // Replace &apos; and &#39; with '
            output = Regex.Replace(output, @"&cent;|&#162;", "¢"); // Replace &cent; and &#162; with ¢
            output = Regex.Replace(output, @"&pound;|&#163;", "£"); // Replace &pound; and &#163; with £
            output = Regex.Replace(output, @"&yen;|&#165;", "¥"); // Replace &yen; and &#165; with ¥
            output = Regex.Replace(output, @"&euro;|&#8364;", "€"); // Replace &euro; and &#8364; with €
            output = Regex.Replace(output, @"&copy;|&#169;", "©"); // Replace &copy; and &#169; with ©
            output = Regex.Replace(output, @"&reg;|&#174;", "®"); // Replace &reg; and &#174; with ®
            output = Regex.Replace(output, @"&trade;|&#8482;", "™"); // Replace &trade; and &#8482; with ™
            return output;
        }
    }
}
