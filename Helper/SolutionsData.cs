using System;
using System.Collections.Generic;
using System.Data;

namespace Retric.WebResourceConsumption.Helper
{
    /// <summary>
    ///     Helper to keep track of solution data
    /// </summary>
    public class SolutionsData
    {
        /// <summary>
        ///     Web types dictionary
        /// </summary>
        public static readonly Dictionary<int, string> WebTypes = new Dictionary<int, string>
        {
            {1, "Webpage (HTML)"},
            {2, "Style Sheet (CSS)"},
            {3, "Script (JScript)"},
            {4, "Data (XML)"},
            {5, "PNG format"},
            {6, "JPG format"},
            {7, "GIF format"},
            {8, "Silverlight (XAP)"},
            {9, "Style Sheet (XSL)"},
            {10, "ICO format"},
            {11, "Vector format (SVG)"},
            {12, "String (RESX)"}
        };

        /// <summary>
        ///     Constructor
        /// </summary>
        public SolutionsData()
        {
            TotalWebResources = new DataTable();
            SolutionList = new Dictionary<Guid, string>();
            SolutionSizes = new Dictionary<Guid, int>();
            OrderedSolutionList = new List<Guid>();

            //Define layout of table
            TotalWebResources.Columns.Add("Guid", typeof(string));
            TotalWebResources.Columns.Add("Name", typeof(string));
            TotalWebResources.Columns.Add("Type", typeof(string));
            TotalWebResources.Columns.Add("Solution", typeof(string));
            TotalWebResources.Columns.Add("Hidden?", typeof(string));
            TotalWebResources.Columns.Add("Size (KB)", typeof(int));
        }


        public Dictionary<Guid, string> SolutionList { get; set; }
        public Dictionary<Guid, int> SolutionSizes { get; set; }
        public List<Guid> OrderedSolutionList { get; set; }
        public DataTable TotalWebResources { get; set; }
    }
}