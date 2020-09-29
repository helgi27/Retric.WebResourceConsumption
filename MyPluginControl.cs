using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Retric.WebResourceConsumption.Helper;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace Retric.WebResourceConsumption
{
    public partial class MyPluginControl : PluginControlBase, IHelpPlugin
    {
        private readonly SolutionsData _sData;
        private Settings _mySettings;

        /// <summary>
        ///     Constructor
        /// </summary>
        public MyPluginControl()
        {
            InitializeComponent();
            _sData = new SolutionsData(); //Init data keeper
        }

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            // Loads or creates the settings for the plugin
            if (!SettingsManager.Instance.TryLoad(GetType(), out _mySettings))
            {
                _mySettings = new Settings();

                LogWarning("Settings not found => a new settings file has been created!");
            }
            else
            {
                LogInfo("Settings found and loaded");
            }
        }

        /// <summary>
        ///     First worker which gets a list of all solutions for future reference after web resources are loaded
        /// </summary>
        private void GetSolutions()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Getting Solutions",
                Work = (worker, args) =>
                {
                    var qex = new QueryExpression("solution")
                    {
                        ColumnSet = new ColumnSet("friendlyname", "solutionid",
                            "uniquename") //Get friendly name, id and unique name
                    };

                    args.Result = Service.RetrieveMultiple(qex);
                },
                PostWorkCallBack = args =>
                {
                    if (args.Error != null)
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    if (args.Result is EntityCollection result)
                        foreach (var item in result.Entities)
                            _sData.SolutionList.Add(item.Id,
                                $"{item.GetAttributeValue<string>("friendlyname")} ({item.GetAttributeValue<string>("uniquename")})");

                    //init size holder for all solutions
                    foreach (var solution in _sData.SolutionList) _sData.SolutionSizes.Add(solution.Key, 0);

                    //Once we have the solutions, lets get the resources -- Will take memory
                    ExecuteMethod(GetWebResources);
                }
            });
        }

        /// <summary>
        ///     The actual worker method to get all web resources
        /// </summary>
        private void GetWebResources()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Getting WebResources",
                Work = (worker, args) =>
                {
                    var resultCollection = PageThroughWebResourcesResultCollection(worker);

                    //The actual return of the results
                    args.Result = resultCollection;
                },
                ProgressChanged = e =>
                {
                    // it will display number of web resources it has gone through already
                    SetWorkingMessage(e.ProgressPercentage.ToString());
                },
                PostWorkCallBack = args =>
                {
                    if (args.Error != null)
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    var totalSize = 0;
                    if (args.Result is EntityCollection result)
                    {
                        totalSize = ProcessResults(result, totalSize);

                        labelInfo.Text =
                            $"Total webresources found: {_sData.TotalWebResources.Rows.Count} ({totalSize / 1024}MB)";
                        _sData.TotalWebResources.DefaultView.Sort = "Size (KB) desc";
                        dataGridView1.DataSource = _sData.TotalWebResources;
                        dataGridView1.Columns[0].Visible = false;
                        dataGridView1.Columns[1].Width = 400;
                    }

                    //Draw basic chart as well
                    ShowChart();
                }
            });
        }

        /// <summary>
        /// Process the actual results
        /// </summary>
        /// <param name="result"></param>
        /// <param name="totalSize"></param>
        /// <returns></returns>
        private int ProcessResults(EntityCollection result, int totalSize)
        {
            //Going through all web resources to calculate size
            foreach (var item in result.Entities)
            {
                var size = GetOriginalLengthInBytes(item.GetAttributeValue<string>("content"));
                var solutionid = item.GetAttributeValue<Guid>("solutionid");

                _sData.TotalWebResources.Rows.Add(item.Id.ToString(),
                    item.GetAttributeValue<string>("name"),
                    SolutionsData.WebTypes[item.GetAttributeValue<OptionSetValue>("webresourcetype").Value],
                    _sData.SolutionList[solutionid],
                    item.GetAttributeValue<BooleanManagedProperty>("ishidden").Value, size / 1024);
                //Add to size as well
                _sData.SolutionSizes[solutionid] += size / 1024; //(Megabytes)
                totalSize += size / 1024; //(Megabytes)
            }

            return totalSize;
        }

        /// <summary>
        /// The main paging method responsible for getting all web resources from org.
        /// </summary>
        /// <param name="worker"></param>
        /// <returns></returns>
        private EntityCollection PageThroughWebResourcesResultCollection(BackgroundWorker worker)
        {
            EntityCollection resultCollection = null;
            var page = 1;
            var queryCount = 200; //Page size
            var totalProcessed = 0;
            var qex = new QueryExpression("webresource")
            {
                ColumnSet = new ColumnSet("name", "webresourcetype", "solutionid", "ishidden", "content"),
                PageInfo = new PagingInfo {Count = queryCount, PageNumber = page}
            };

            while (true)
            {
                var tmpResult = Service.RetrieveMultiple(qex);
                if (resultCollection == null)
                {
                    resultCollection = tmpResult;
                }
                else
                {
                    resultCollection.Entities.AddRange(tmpResult.Entities);
                    resultCollection.MoreRecords = tmpResult.MoreRecords;
                    resultCollection.PagingCookie = tmpResult.PagingCookie;
                    resultCollection.TotalRecordCount = tmpResult.TotalRecordCount;
                    resultCollection.TotalRecordCountLimitExceeded = tmpResult.TotalRecordCountLimitExceeded;
                }

                // Check for more records, if it returns true.
                if (resultCollection.MoreRecords)
                {
                    totalProcessed += queryCount;
                    worker.ReportProgress(totalProcessed);
                    // Increment the page number to retrieve the next page.
                    qex.PageInfo.PageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.
                    qex.PageInfo.PagingCookie = resultCollection.PagingCookie;
                }
                else
                {
                    // If no more records are in the result nodes, exit the loop.
                    break;
                }
            } //while ends

            return resultCollection;
        }

        /// <summary>
        ///     Draw the basic chart
        /// </summary>
        private void ShowChart()
        {
            chart1.Series.Add("Consumption");
            var sortedDict = from entry in _sData.SolutionSizes orderby entry.Value descending select entry;
            _sData.OrderedSolutionList = sortedDict.Select(x => x.Key).ToList();
            foreach (var solution in sortedDict.Take(5))
                chart1.Series["Consumption"].Points.AddXY(_sData.SolutionList[solution.Key], solution.Value / 1024);
            //chart title  
            chart1.Titles.Add("Largest solutions (MB)");
        }

        #region Button clicks&events

        /// <summary>
        ///     The click event of the main button "Load webresources"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BtnLoadAllWebResources(object sender, EventArgs e)
        {
            ExecuteMethod(GetSolutions);
        }

        /// <summary>
        ///     When one specific chart column is clicked we filter the grid on that solution
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Chart_mouseDown(object sender, MouseEventArgs e)
        {
            LogInfo("Click chart");
            var r = chart1.HitTest(e.X, e.Y);

            if (r.ChartElementType == ChartElementType.DataPoint)
            {
                var index = r.PointIndex;
                _sData.TotalWebResources.DefaultView.RowFilter =
                    $"[Solution] LIKE '%{_sData.SolutionList[_sData.OrderedSolutionList[index]]}%'";
            }
            else
            {
                _sData.TotalWebResources.DefaultView.RowFilter = null;
            }
        }

        /// <summary>
        ///     Close tool button click
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }

        /// <summary>
        ///     This event occurs when the plugin is closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyPluginControl_OnCloseTool(object sender, EventArgs e)
        {
            // Before leaving, save the settings
            SettingsManager.Instance.Save(GetType(), _mySettings);
        }

        /// <summary>
        ///     This event occurs when the connection has been updated in XrmToolBox
        /// </summary>
        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail,
            string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);

            if (_mySettings != null && detail != null)
            {
                _mySettings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl;
                LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);
            }
        }

        #endregion

        #region Helper methods

        /// <summary>
        ///     Helper to get byte size from base64 string, method from web.
        /// </summary>
        /// <param name="base64String"></param>
        /// <returns></returns>
        private int GetOriginalLengthInBytes(string base64String)
        {
            if (string.IsNullOrEmpty(base64String)) return 0;

            var characterCount = base64String.Length;
            var paddingCount = base64String.Substring(characterCount - 2, 2)
                .Count(c => c == '=');
            return 3 * (characterCount / 4) - paddingCount;
        }

        public string HelpUrl => "https://github.com/helgi27/Retric.WebResourceConsumption/wiki";

        #endregion
    }
}