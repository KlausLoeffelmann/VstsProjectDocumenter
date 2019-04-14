using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.WebApi;
using VstsProjectDocumenter.DataStructures;
using WorkItem = Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models.WorkItem;

namespace VstsProjectDocumenter
{
    public partial class ProjectDocumenterRibbon
    {
        private async void DocumentProjectbutton_Click(object sender, RibbonControlEventArgs e)
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
                ProjectHttpClient projClient = await vsConn.GetClientAsync<ProjectHttpClient>();
                var projects = projClient.GetProjects().Result;

                var adopIteration = AdopIteration.GetIterations(projects.Where(pro => pro.Name == projectName).
                    Select((item) => workItemTracking.GetClassificationNodeAsync(
                        project: item.Name,
                        structureGroup: TreeStructureGroup.Iterations,
                        depth: 3).Result));

                foreach (var iterationItem in adopIteration.
                    Where((item) => item.Level == 2).
                    OrderBy((item) => item.StartDate))
                {

                    InsertIterationAsHeadline(iterationItem);
                    var iterationTable = BeginNewIterationTable();

                    Wiql query = new Wiql()
                    {
                        Query = "SELECT [Id], [Title], [State] FROM workitems" +
                                " WHERE ([Work Item Type] = 'Bug'" +
                                " OR [Work Item Type] = 'Product Backlog Item'" +
                                " OR [Work Item Type] = 'Task')" +
                                $" AND [Iteration Path] = '{iterationItem.FullPath}'"
                    };

                    WorkItemQueryResult queryResults;
                    try
                    {
                        queryResults = await witClient.QueryByWiqlAsync(query, projectName);
                    }
                    catch (Exception ex)
                    {
                        throw (ex);
                    }

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
                            string[] fieldNameArray;
                            try
                            {
                                fieldNameArray = new string[]{
                                    $"System.{nameof(CoreField.Id)}",
                                    $"System.{nameof(CoreField.IterationPath)}",
                                    $"System.{nameof(CoreField.WorkItemType)}",
                                    $"System.{nameof(CoreField.Title)}",
                                    $"System.{nameof(CoreField.Description)}",
                                    $"System.{nameof(CoreField.AssignedTo)}",
                                    $"System.{nameof(CoreField.State)}",
                                    $"System.{nameof(CoreField.IsDeleted)}"};

                                var workItem = await witClient.GetWorkItemAsync(item.Id, fieldNameArray);
                                AddWorkItemRow(iterationTable, workItem);
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }

                    ThisAddIn.ThisApplication.Selection.EndOf(WdUnits.wdStory);
                    ThisAddIn.ThisApplication.Selection.InsertNewPage();
                }


                //var ad = ThisAddIn.ThisApplication.ActiveDocument;
                //var newTable = ad.Tables.Add(ThisAddIn.ThisApplication.Selection.Range,
                //                1, 3);
                //var cell = newTable.Cell(1, 1);
                //cell.Range.Text = "Iteration/Leistungspaket";
                //cell = newTable.Cell(1, 2);
                //cell.Range.Text = "Start-Datum";
                //cell = newTable.Cell(1, 3);
                //cell.Range.Text = "End-Datum";

                //foreach (var iterationItem in adopIteration.
                //    Where((item)=> item.Level==2).
                //    OrderBy((item)=>item.StartDate))
                //{
                //    var currentRow = newTable.Rows.Add();
                //    currentRow.Cells[1].Range.Text = iterationItem.Name;
                //    currentRow.Cells[2].Range.Text = iterationItem.StartDate.ToShortDateString();
                //    currentRow.Cells[3].Range.Text = iterationItem.FinishDate.ToShortDateString();
                //}
            }
        }

        private void AddWorkItemRow(Table iterationTable, WorkItem workItem)
        {
            Row currentRow;

            try
            {
                currentRow = iterationTable.Rows.Add();
                currentRow.Range.Font.Size = 7;
                currentRow.Range.Font.Bold = 0;
                currentRow.Cells[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                currentRow.Cells[2].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                currentRow.Cells[3].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                currentRow.Cells[4].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                currentRow.Cells[1].Range.Text = TryGetFieldContent(workItem, $"System.{nameof(CoreField.Title)}", "Kein Titel definiert.");
                currentRow.Cells[2].Range.Text = TryGetFieldContent(workItem, $"System.{nameof(CoreField.AssignedTo)}", "Team effort.");
                currentRow.Cells[3].Range.Text = TryGetFieldContent(workItem, $"System.{nameof(CoreField.WorkItemType)}", "Allgemein.");
                TryInsertHtmlAtRange(currentRow.Cells[4].Range, TryGetFieldContent(workItem, $"System.{nameof(CoreField.Description)}", "<p>Keine nähere Beschreibung hinterlegt.</p>"));
                currentRow.Cells[1].Range.Font.Bold = 1;
                currentRow.Cells[4].Range.Font.Size = 7;
            }
            catch (Exception ex)
            {
            }
        }

        string TryGetFieldContent(WorkItem workItem, string fieldName, string defaultText)
        {
            if (workItem.Fields.ContainsKey(fieldName))
            {
                return (string)workItem.Fields[fieldName];
            }
            else
            {
                return defaultText;
            }
        }

        void TryInsertHtmlAtRange(Range range, string html)
        {
            try
            {
                html = html.Replace("ä", "&auml").
                            Replace("Ä", "&Auml").
                            Replace("ü", "&uuml").
                            Replace("Ü", "&Uuml").
                            Replace("ö", "&ouml").
                            Replace("Ö", "&Ouml").
                            Replace("ß", "&szlig");

                // If you see Trust messages for the file location while running the add-in,
                // add this path to the trusted locations.
                // e.g. C:\Users\[username]\AppData\Local\Temp\
                string path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ProjectDocumenter\\";
                path += Guid.NewGuid().ToString("D") + ".tmp";
                ASCIIEncoding utf8 = new ASCIIEncoding();
                System.IO.File.WriteAllBytes(path, utf8.GetBytes("<html>" + html));
                range.InsertFile(path, ConfirmConversions: false);
                System.IO.File.Delete(path);
            }
            catch (Exception)
            {
            }
        }


        void InsertIterationAsHeadline(AdopIteration iteration)
        {
            ThisAddIn.ThisApplication.Selection.TypeText(
                $"Leistungspaket: {iteration.Name} vom {iteration.StartDate:dd.MM.yyyy} bis {iteration.FinishDate:dd.MM.yyyy}");
            object codeStyle = "Heading 1";
            var range = ThisAddIn.ThisApplication.Selection.Range;
            range.set_Style(ref codeStyle);
            ThisAddIn.ThisApplication.Selection.TypeParagraph();
            ThisAddIn.ThisApplication.Selection.TypeParagraph();
            //ThisAddIn.ThisApplication.Selection.TypeText("Hier geht es weiter.");
            //ThisAddIn.ThisApplication.Selection.TypeParagraph();
            //ThisAddIn.ThisApplication.Selection.TypeText("Hier geht es noch weiter.");
            //ThisAddIn.ThisApplication.Selection.TypeParagraph();
            //range = ThisAddIn.ThisApplication.Selection.Range;
            //range.Text = "Test";
        }

        Table BeginNewIterationTable()
        {
            var ad = ThisAddIn.ThisApplication.ActiveDocument;
            var newTable = ad.Tables.Add(ThisAddIn.ThisApplication.Selection.Range,
                            1, 4);
            newTable.Columns.AutoFit();

            newTable.Columns[1].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            newTable.Columns[2].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            newTable.Columns[3].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            newTable.Columns[4].PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;

            newTable.Columns[1].PreferredWidth = 15;
            newTable.Columns[2].PreferredWidth = 18;
            newTable.Columns[3].PreferredWidth = 12;
            newTable.Columns[4].PreferredWidth = 55;

            var cell = newTable.Cell(1, 1);
            cell.Range.Text = "Titel";
            cell = newTable.Cell(1, 2);
            cell.Range.Text = "Zuständiger Entwickler";
            cell = newTable.Cell(1, 3);
            cell.Range.Text = "Arbeitselemente-Typ";
            cell = newTable.Cell(1, 4);
            cell.Range.Text = "Beschreibung";
            newTable.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            newTable.Rows[1].Range.Font.Size = 9;
            newTable.Rows[1].Range.Font.Bold = 1;
            return newTable;
        }
    }
}