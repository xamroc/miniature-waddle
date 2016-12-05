using System;
using System.Text.RegularExpressions;
using SpreadsheetGear;
using Newtonsoft.Json;

namespace PRoschke_ExcelTest3
{
    /// <summary>
    /// Class for holding and managing the in-memory Excel files and spreadsheet gear workbook instance.
    /// </summary>
    public class XOutput
    {
        private SpreadsheetGear.IName sgOName;
        private SpreadsheetGear.IRange topLeftCell;
        private string name;
        private long cellCount = 0;
        private string description = "";
        private string fgColor = "";
        private string bgColor = "";
        private string type = "";

        [JsonIgnore]
        public IName SgOName
        {
            get
            {
                return sgOName;
            }
        }
        [JsonIgnore]
        public IRange TopLeftCell
        {
            get
            {
                return topLeftCell;
            }
        }

        public string Name
        {
            get
            {
                return name;
            }
        }
        [JsonIgnore]
        public long CellCount
        {
            get
            {
                return cellCount;
            }
        }
        /// <summary>
        /// Comments assoicated with the output.  Automatically populated from Name or range comments 
        /// but only for XLS files (others not presently supported - so maybe use set and pull from Xcomment_[filed] names)
        /// </summary>
        [JsonIgnore]
        public string Description
        {
            get
            {
                return description;
            }
            set
            {
                description = value;
            }
        }

        /// <summary>
        /// Returns the top left cells font colour in ARGB hex format.
        /// </summary>
        [JsonIgnore]
        public string FgColor
        {
            get
            {
                return fgColor;
            }
        }

        /// <summary>
        /// Returns the top left cells background colour in ARGB hex format.
        /// </summary>
        [JsonIgnore]
        public string BgColor
        {
            get
            {
                return bgColor;
            }
        }


        public string Value
        {
            get
            {
                if (cellCount == 1) return topLeftCell.Text;

                // we have a range that was specified.
                if (sgOName.Name.StartsWith("XoutputTable_"))
                {  // We have some header rows so we can make a nicer objects
                    //return JsonConvert.SerializeObject(Helper.GetTableFromIRange(sgOName.RefersToRange));
                    return JsonConvert.SerializeObject(sgOName.RefersToRange.GetDataTable(SpreadsheetGear.Data.GetDataFlags.None));
                } else
                //return JsonConvert.SerializeObject(Helper.GetArrayFromIRange(sgOName.RefersToRange));
                return JsonConvert.SerializeObject(sgOName.RefersToRange.GetDataTable(SpreadsheetGear.Data.GetDataFlags.NoColumnHeaders));
            }
            set
            {
                topLeftCell.Formula = value;
            }
        }


        public XOutput (SpreadsheetGear.IName sgout)
        {
            sgOName = sgout;
            cellCount = sgOName.RefersToRange.CellCount;  // TODO Names can refer to ranges and actually groups of ranges... support for that to be added later for XInputs.
            topLeftCell = sgout.RefersToRange[0, 0];
            //Strip out Xoutput_ identifier from Name..
            name = sgOName.Name.StartsWith("Xoutput_") ? sgOName.Name.Substring(8) : sgOName.Name;

            //Strip out XoutputTable_ identifier from Name..
            name = sgOName.Name.StartsWith("XoutputTable_") ? sgOName.Name.Substring(13) : Name;


            // Check for comment Filds .... these can exist in a number of areas so let's check
            // Also note that SG does not support Comments for .xlsx .xlsm files ... only the older XLS files.  So people could save down.
            if (sgOName.Comment.Length > 0) description = sgOName.Comment;
            else if (topLeftCell.Comment != null) description = topLeftCell.Comment.ToString();


            // But override description if a Xdescription_ field exists  -- needed for XLSX files
            if (topLeftCell.Worksheet.Workbook.Names["Xdescription_" + name] != null) description = topLeftCell.Worksheet.Workbook.Names["Xdescription_" + name].RefersToRange.Text;


            bgColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(topLeftCell.Interior.Color.ToArgb()));
            fgColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(topLeftCell.Font.Color.ToArgb()));

        }

    }
}