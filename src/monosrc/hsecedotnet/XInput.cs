using System;
using System.Text.RegularExpressions;
using SpreadsheetGear;
using Newtonsoft.Json;

namespace PRoschke_ExcelTest3
{
    /// <summary>
    /// Class for holding and managing the in-memory Excel files and spreadsheet gear workbook instance.
    /// </summary>
    public class XInput
    {
		[JsonIgnore]
        private SpreadsheetGear.IName sgIName;

		[JsonIgnore]
        private SpreadsheetGear.IRange topLeftCell;

        private string name;
        private long cellCount = 0;
        private string description = "";
        private string fgColor = "";
        private string bgColor = "";
        private string valueList = "";
        private string valType = "";
        private string dataType = "";
        private bool hasValidation= false;
        private string defaultValue = "";

		[JsonIgnore]
        private SpreadsheetGear.IValidation validation=null;

        [JsonIgnore]
        public IName SgIName
        {
            get
            {
                return sgIName;
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

    	[JsonProperty(PropertyName = "Name")]
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
        /// Description assoicated with the input.  Can explicity be set by defining a named call: Xdescription_[name].
        /// Otherwise automatically populated from Name or range comments but only for XLS files (others not presently supported)
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
        [JsonProperty(PropertyName = "FgColor")]
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
        [JsonProperty(PropertyName = "BgColor")]
        public string BgColor
        {
            get
            {
                return bgColor;
            }
       }
        /// <summary>
        /// Get's a best guess of the permitted values as a comma separated list.  This is not that easy to do 
        /// and this property shoudl only be used for demos.  For prodution list based validations should rely 
        /// on the native validation capatilities already embedded in this class from spreadsheet gear (you just
        /// don't get to see the list).  
        /// </summary>
        [JsonIgnore]
        public string ValueList
        {
            get
            {
                return valueList;
            }
        }

        [JsonProperty(PropertyName = "ValueType")]
        public string ValType
        {
            get
            {
                return valType;
            }
        }

        [JsonIgnore]
        public bool HasValidation
        {
            get
            {
                return hasValidation;
            }
        }

        [JsonIgnore]
        public IValidation Validation
        {
            get
            {
                return validation;
            }
        }
        [JsonIgnore]
        public bool IsValid
        {
            get
            {
                if (validation != null) return validation.Value;
                else return true;
            }
        }
        [JsonIgnore]
        public String ErrorMessage
        {
            get
            {
                string str="";
                if (!IsValid)   // then concatenate whatever info we can find in the error definiting
                {
                    if (validation.ErrorTitle!=null)
                    {
                        str = validation.ErrorTitle;
                    }
                    if (validation.ErrorMessage != null)
                    {
                        if (str.Length > 0) str = str + " | ";
                        str = str + validation.ErrorMessage;
                    }
                    if (str.Length == 0) str = "Validation " + validation.AlertStyle.ToString()+ " on input " + name + ".";
                }
                return str;
            }
        }

    	[JsonProperty(PropertyName = "Value")]
        public string Value
        {
            get
            {
                return topLeftCell.Text;
            }
            set
            {
                topLeftCell.Value = value;
            }
        }

        [JsonProperty(PropertyName = "DataType")]
        public string DataType
        {
            get
            {
                return dataType;
            }

        }

		[JsonProperty(PropertyName = "HTMLDataType")]
        public string HTMLDataType
        {
            get
            {
                if (dataType.Equals("Number")) return "number";
                if (dataType.Equals("Date")) return "date";
                if (dataType.Equals("Logical")) return "checkbox";
                if (dataType.Equals("Datetime")) return "datetime";
                if (dataType.Equals("Percent")) return "number";
                if (dataType.Equals("Time")) return "time";
                return "text";
            }

        }

        public void resetToDefaultValue()
        {
            topLeftCell.Value = defaultValue;
        }

        public XInput(SpreadsheetGear.IName sgin)
        {
            sgIName = sgin;
            cellCount = sgIName.RefersToRange.CellCount;  // TODO Names can refer to ranges and actually groups of ranges... support for that to be added later for XInputs.
            topLeftCell = sgin.RefersToRange[0, 0];
            defaultValue = topLeftCell.Value==null ? "" : topLeftCell.Value.ToString();

            //Strip out XInput_ identifier from Name..
            name = sgIName.Name.StartsWith("Xinput_") ? sgIName.Name.Substring(7) : sgIName.Name;

            // Get data type informaiton for input
            dataType = topLeftCell.ValueType.ToString();
            if (dataType.Equals(SpreadsheetGear.ValueType.Number.ToString())) dataType = topLeftCell.NumberFormatType.ToString();
            // it's possible that the number formattype is "None" .. in which case revert back to "Number"
            if (dataType.Equals(SpreadsheetGear.NumberFormatType.None.ToString())) dataType = SpreadsheetGear.NumberFormatType.Number.ToString();
            // it's possible that the number formattype is "Genera" .. in which case revert back to "Number"
            if (dataType.Equals(SpreadsheetGear.NumberFormatType.General.ToString())) dataType = SpreadsheetGear.NumberFormatType.Number.ToString();
            // Local date formats are terrible - we will change all date formats to the standard ISO format (Also matches HTML 5).  Does not affect calcs just mapping things in and out of strings.
            if (dataType.Equals(SpreadsheetGear.NumberFormatType.Date.ToString())) topLeftCell.NumberFormat = "yyyy-mm-dd";


            // SET DESCRIPTION member
            // Check for comment Filds .... these can exist in a number of areas so let's check
            // Also note that SG does not support Comments for .xlsx .xlsm files ... only the older XLS files.  So people could save down.
            if (sgIName.Comment.Length > 0) description = sgIName.Comment;
            else if (topLeftCell.Comment != null) description = topLeftCell.Comment.ToString();

            // But override description if a Xdescription_ field exists
            if (topLeftCell.Worksheet.Workbook.Names["Xdescription_" + name] != null) description = topLeftCell.Worksheet.Workbook.Names["Xdescription_" + name].RefersToRange.Text;

            // Set up VAlidation rule properties
            if (topLeftCell.HasValidation)
            {
                if (topLeftCell.Validation.Type == SpreadsheetGear.ValidationType.List)
                {
                    valueList = topLeftCell.Validation.Formula1;   //could be regular comma separated list
                    if(valueList.StartsWith("="))
                    {
                        //we are dealing with a range .... probably.  I guess other things could be used here like lists of ranges or formulas that return ranges. 
                        //we'll just to the basics here.  Correct validation will happen anyway via Spreadsheet gear.

                        string vref = valueList.Substring(from: "=");   // we have some strimg manipulation coming up so I grabbed a helper function
                        valueList = "";
                        IWorksheet ws = topLeftCell.Worksheet;
                        if (vref.Contains("!"))
                        {
                            //And the range is on a different worksheet.
                            ws=ws.Workbook.Worksheets[vref.Substring(until: "!")];
                            vref = vref.Substring(from: "!");
                        }
                        // select cells, save to byte buffer and convert to string... works but has uneven results - so skipping next line and using for loop below instead
                        //valueList = System.Text.Encoding.Default.GetString(ws.Cells[valueList].SaveToMemory(SpreadsheetGear.FileFormat.CSV));
                        foreach (IRange ir in ws.Cells[vref])
                        {
                            valueList += ir.Value + ", ";
                        }
                        valueList = valueList.Substring(0, valueList.Length - 2);  //remove last comma and space
                    }

                }
                hasValidation = true;
                validation =topLeftCell.Validation;
                valType = validation.Type.ToString();
            }

            bgColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(topLeftCell.Interior.Color.ToArgb()));
            fgColor = System.Drawing.ColorTranslator.ToHtml(System.Drawing.Color.FromArgb(topLeftCell.Font.Color.ToArgb()));
        }

    }
}