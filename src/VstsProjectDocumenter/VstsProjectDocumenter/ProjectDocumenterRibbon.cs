using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.WebApi;
using VstsProjectDocumenter.DataStructures;

namespace VstsProjectDocumenter
{
    public partial class ProjectDocumenterRibbon
    {
        private void DocumentProjectbutton_Click(object sender, RibbonControlEventArgs e)
        {
            string projectUrl = Properties.Settings.Default.ProjectUrl;
            string projectName = Properties.Settings.Default.ProjectName;

            if (string.IsNullOrWhiteSpace(projectUrl) || string.IsNullOrWhiteSpace(projectName))
            {
                if (new AddProjectUrlForm().ShowDialog(out projectUrl, out projectName) == DialogResult.OK)
                {
                    Properties.Settings.Default.ProjectUrl = projectUrl;
                    Properties.Settings.Default.ProjectName = projectName;
                    Properties.Settings.Default.Save();
                }
            }

            if (!string.IsNullOrWhiteSpace(projectUrl))
            {
                VssConnection vsConn = new VssConnection(new Uri(projectUrl), new VssClientCredentials());
                WorkItemTrackingHttpClient witClient;

                try
                {
                    //create http client and query for resutls
                    witClient = vsConn.GetClient<WorkItemTrackingHttpClient>();
                }
                catch (Exception ex)
                {
                    throw (ex);
                }

                var workItemTracking = vsConn.GetClient<WorkItemTrackingHttpClient>();
                ProjectHttpClient projClient = vsConn.GetClientAsync<ProjectHttpClient>().Result;
                var projects = projClient.GetProjects().Result;

                var adopIteration = AdopIteration.GetIterations(projects.Where(pro => pro.Name == projectName).
                    Select((item) => workItemTracking.GetClassificationNodeAsync(
                        project: item.Name,
                        structureGroup: TreeStructureGroup.Iterations,
                        depth: 3).Result));

                var ad = ThisAddIn.ThisApplication.ActiveDocument;
                var newTable = ad.Tables.Add(ThisAddIn.ThisApplication.Selection.Range,
                                1, 3);
                var cell = newTable.Cell(1, 1);
                cell.Range.Text = "Iteration/Leistungspaket";
                cell = newTable.Cell(1, 2);
                cell.Range.Text = "Start-Datum";
                cell = newTable.Cell(1, 3);
                cell.Range.Text = "End-Datum";

                foreach (var iterationItem in adopIteration.
                    Where((item)=> item.Level==2).
                    OrderBy((item)=>item.StartDate))
                {
                    var currentRow = newTable.Rows.Add();
                    currentRow.Cells[1].Range.Text = iterationItem.Name;
                    currentRow.Cells[2].Range.Text = iterationItem.StartDate.ToShortDateString();
                    currentRow.Cells[3].Range.Text = iterationItem.FinishDate.ToShortDateString();
                }
            }
        }

        void QueryWorkItems(WorkItemTrackingHttpClient witClient)
        {
            Wiql query = new Wiql() { Query = "SELECT [Id], [Title], [State] FROM workitems WHERE [Work Item Type] = 'Bug' AND [Assigned To] = @Me" };
            WorkItemQueryResult queryResults = witClient.QueryByWiqlAsync(query).Result;

            //Display results in console
            if (queryResults == null || queryResults.WorkItems.Count() == 0)
            {
                Console.WriteLine("Query did not find any results");
            }
            else
            {
                foreach (var item in queryResults.WorkItems)
                {
                    Console.WriteLine(item.Id);
                }
            }
        }

        static void GetIterations(WorkItemClassificationNode currentIteration)
        {
            Console.WriteLine(currentIteration.Name);
            if (currentIteration.Children != null)
            {
                foreach (var ci in currentIteration.Children)
                {
                    GetIterations(ci);
                }
            }
        }
    }
}
