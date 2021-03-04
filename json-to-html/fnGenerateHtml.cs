using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using HtmlAgilityPack;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace json_to_html
{
    public static class fnGenerateHtml
    {
        [FunctionName("fnGenerateHtml")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            // start tracing
            var traceID = Guid.NewGuid().ToString();
            log.LogInformation($"{traceID} - fnGenerateHtml");

            // get body JSON and process
            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            JObject jsonData = JObject.Parse(requestBody);
            var htmlResult = GenerateHtmlOutput(jsonData);

            // return the json result
            return new OkObjectResult(htmlResult);
        }

        private static string GenerateHtmlOutput(JObject jsonData)
        {
            // TODO: Change hard-coded template below to use a different approach, perhaps a root property of the JSON
            var outputTemplate = "";

            // intialise variables
            var htmlOutput = "<html><head></head><style>OL { counter-reset: item } LI { display: block } LI:before { content: counters(item, \".\") \" \"; counter-increment: item }</style><body>";
            string[] keyPhrases = new string[] { "risk appetite", "key risk indicators", "controls" };
            string subHeadingStart = "Sub Risk Type";

            dynamic data = JsonConvert.DeserializeObject(jsonData.ToString());

            // get all the tables in the json
            var tableIndex = 0;
            foreach (var table in data.tables)
            {
                var tableHtml = $"<p><ol style=\"counter-reset: item {tableIndex}\">";
                var rowIndex = 0;
                var title = String.Empty;
                var lastHeading = string.Empty;
                var liTierCount = 0; // count the number of <li> tags that are not clsoed
                var subRiskTypeCellCount = 0;
                foreach (var curRow in table.rows)
                {
                    var rowHtml = string.Empty;

                    // first row is the name and description
                    if (rowIndex == 0)
                    {
                        var lines = curRow.cells[0].lines;
                        title = lines[0].text;
                        string description = lines[1].text;
                        rowHtml = $"<li>{title}</li>{title} is {description.Replace(title + " is ", "")}<ol>";
                    }
                    // other rows need to be processed seperately
                    else
                    {
                        // get the text of the first cell, as this can indicate a specific type of cell
                        string firstCellContent = curRow.cells[0].lines[0].text;

                        // if this a "key phrase" row, take the first cell and the last (sometimes there is no content or nonsense in-between)
                        if (!String.IsNullOrEmpty(firstCellContent) && keyPhrases.Contains(firstCellContent.Trim().ToLower()))
                        {
                            // close the extra tiers, if necessary
                            for (int i = liTierCount; i > 0; i--)
                            {
                                rowHtml += "</ol></li>";
                                liTierCount--;
                            }

                            // assume that the key phrase text is in the second cell
                            // if, however, there are more than 2 cells in this row and the last cell is not blank, take the third cell
                            var keyPhraseLines = curRow.cells[1].lines;
                            if (curRow.cells.Count > 2 && String.IsNullOrWhiteSpace(curRow.cells[curRow.cells.Count - 1].lines[0].text.ToString()) == false)
                            {
                                keyPhraseLines = curRow.cells[curRow.cells.Count - 1].lines;
                            }
                            rowHtml += $"<li>{firstCellContent}<br/>{GetAllText(keyPhraseLines)}</li>";
                        }
                        else if (firstCellContent.StartsWith(subHeadingStart))
                        {
                            rowHtml = "<li>Sub Risk Types<ol>";
                            liTierCount++;
                            subRiskTypeCellCount = curRow.cells.Count;
                        }
                        // other rows are not key phrases
                        else
                        {
                            // for these rows, we create a template to use, based on the number of cells, and whether the first cell indicates a new section or not
                            var newSection = false;
                            var template = String.Empty;

                            // check if we are dealing with a new sub risk type
                            if (firstCellContent != lastHeading && String.IsNullOrEmpty(firstCellContent.Trim()) == false)
                            {
                                newSection = true;
                                lastHeading = firstCellContent;
                            }

                            // define the template to use, based on the number of cells for the sub risk type row we are on
                            if (subRiskTypeCellCount == 4)
                            {
                                if (newSection)
                                {
                                    template =
                                        "<li>|subRiskType1|" +
                                            "<ol>" +
                                                "<li>|subRiskType2|" +
                                                    "<ol>" +
                                                        "<li>Description<br/>|description|</li>" +
                                                        "<li>Relation to other risks / activities<br/>|relation|</li>" +
                                                    "</ol>" +
                                                "</li>";

                                    // close the previous section, if required
                                    if (liTierCount == 2)
                                    {
                                        template =
                                            "</ol>" +
                                        "</li>" + template;
                                        liTierCount--;
                                    }

                                    // increment the tier count for the open <li> for this new section
                                    liTierCount++;
                                }
                                else
                                {
                                    template =
                                                "<li>|subRiskType2|" +
                                                    "<ol>" +
                                                        "<li>Description<br/>|description|</li>" +
                                                        "<li>Relation to other risks / activities<br/>|relation|</li>" +
                                                    "</ol>" +
                                                "</li>";
                                }
                            }
                            else
                            {
                                template =
                                    "<li>|subRiskType1|" +
                                        "<ol>" +
                                            "<li>Description<br/>|description|</li>" +
                                            "<li>Relation to other risks / activities<br/>|relation|</li>" +
                                        "</ol>" +
                                    "</li>";
                            }

                            // add the 2nd-last and last cells
                            rowHtml = template
                                .Replace("|subRiskType1|", firstCellContent)
                                .Replace("|subRiskType2|", curRow.cells[1].lines[0].text.ToString())
                                .Replace("|description|", GetAllText(curRow.cells[curRow.cells.Count - 2].lines))
                                .Replace("|relation|", GetAllText(curRow.cells[curRow.cells.Count - 1].lines));
                        }
                    }

                    // increment row Index
                    tableHtml += rowHtml;
                    rowIndex++;
                }

                // close current table and increment table index
                tableHtml += "</ol></ol></p>";
                htmlOutput += tableHtml;
                tableIndex++;
            }

            // close html output and return
            htmlOutput += "</body></html>";
            return htmlOutput;
        }

        private static string GetAllText(dynamic lines)
        {
            string returnText = String.Empty;
            foreach (var line in lines)
            {
                returnText += HttpUtility.HtmlEncode(line.text) + "<br />";
            }
            return returnText;
        }
    }
}
