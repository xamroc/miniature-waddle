using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using SpreadsheetGear;
using Newtonsoft.Json.Linq;
using Microsoft.CSharp;

namespace PRoschke_ExcelTest3
{
    /// <summary>
    /// Class for holding and managing the in-memory Excel files and spreadsheet gear workbook instance.
    /// </summary>
    public class XLObj : IEquatable<XLObj>
    {
        private byte[] data;
        private string md5hash= "";
        private SpreadsheetGear.IWorkbook workbook;
        private long memUsed = 0;
        private bool wasLoadedFromFile = false;
        private int length = 0;
        private double lastCalctime = 0;
        private string mimeType = "application/vnd.ms-excel";

        //Values extracted from One-of named ranges at load time (ie. these should never change during usage)
        private string xname ="";
        private string xrevision ="";
        private string xauthor ="";
        private string xreleaseDate = "";
        private string xnotes="";
        private string xtags="";
        private string debugMessage = "";  //just used to held diagnose parsing errors.
        private bool xdoNotLog=false;
        private bool xdynamicTestUI=true;

        // These Valus are dynamic and should change during usage... should be reset after earch Calculate funtion
        // TODO  move calculation execution a function of this class
        private string xerrorMessage="";
        private string xlogMessage="";

        private Dictionary<string, XInput> dictOfInputs = new Dictionary<string, XInput>();
        private Dictionary<string, XOutput> dictOfOutputs = new Dictionary<string, XOutput>();

        public override string ToString()
        {
            return "Engine Hash: " + md5hash.ToString();
        }

        public override bool Equals(object obj)
        {
            return (this==obj);
        }

        public bool Equals(XLObj other)
        {
            if (other == null) return false;
            return (this.md5hash.Equals(other.md5hash));
        }

        public static bool operator !=(XLObj a, XLObj b)
        {
 
            return !(a==b);
        }

        public static bool operator ==(XLObj a, XLObj b)
        {
            // If left hand side is null...
            if (System.Object.ReferenceEquals(a, null))
            {
                // ...and right hand side is null...
                if (System.Object.ReferenceEquals(b, null))
                {
                    //...both are null and are Equal.
                    return true;
                }

                // ...right hand side is not null, therefore not Equal.
                return false;
            }
            //don't try to compare if b is null
            if (System.Object.ReferenceEquals(b, null)) return false;
            
            // Return true if the fields match:
                return a.md5hash == b.md5hash;
        }
        public override int GetHashCode()
        {
            return md5hash.GetHashCode();  // Not ideal, but since the interface requies and INT not much choice.  Collision probability is low.
        }
        public string MimeType
        {
            get
            {
                return mimeType;
            }
            set
            {
                mimeType=value;
            }
        }

        public bool WasLoadedFromFile
        {
            get
            {
                return wasLoadedFromFile;
            }
        }

        // Can be used to get and set the raw binary Excel file data.
        // Setting also creates correspond SpreadsheetGear object and sets a number of properties based on the
        // insepction of the Excel contents
        public byte[] Data
        {
            get
            {
                return data;
            }
            set
            {
                memUsed = GC.GetTotalMemory(true);
                md5hash = Helper.GetMD5Hash(value);
                data = value;
                // Create a new empty workbook set.
                SpreadsheetGear.IWorkbookSet workbookSet = SpreadsheetGear.Factory.GetWorkbookSet();
                // Open the saved workbook from memory.
                workbook = workbookSet.Workbooks.OpenFromMemory(data);
                workbook.WorkbookSet.Calculation = SpreadsheetGear.Calculation.Manual; 

                //Inspect Named ranges in Excel file
                bool didNotMatch = false;
                bool noXinputs = true;
                foreach (IName nr in workbook.Names ){
                    debugMessage += "Found iName: " + nr.Name + " | ";
                    if (nr.RefersToRange!= null) { // you can delete cells but still have the refernce there.
                    switch (nr.Name)
                    {
                        case "Xrevision":
                            xrevision = nr.RefersToRange.Text;
                            break;
                        case "Xname":
                            xname = nr.RefersToRange.Text;
                            break;
                        case "Xauthor":
                            xauthor = nr.RefersToRange.Text;
                            break;
                        case "XreleaseDate":
                            xreleaseDate= nr.RefersToRange.Text;
                            break;
                        case "Xnotes":
                            xnotes = nr.RefersToRange.Text;
                            break;
                        case "XdoNotLog":
                            xdoNotLog = !nr.RefersToRange.Text.Equals("TRUE");
                            break;
                        case "XdynamicTestUI":
                            xdoNotLog = nr.RefersToRange.Text.Equals("TRUE");
                            break;
                        case "XerrorMessage":  // needs to also be updated after each calc
                            xerrorMessage = nr.RefersToRange.Text;
                            break;
                        case "Xtags":
                            xtags = nr.RefersToRange.Text;
                            break;
                        case "XlogMessage":  // needs to also be updated after each calc
                            xlogMessage = nr.RefersToRange.Text;
                            break;
                        default:
                            if(nr.Name.StartsWith("Xinput_"))
                            {
                                XInput xin = new XInput(nr);
                                // Add to dictionary but strip out Xinput_ prefix in key
                                dictOfInputs.Add(xin.Name, xin);  
                            } else if (nr.Name.StartsWith("Xoutput_") || nr.Name.StartsWith("XoutputTable_"))
                            {
                                // BUG: In spreadhseet gear when a named range refers to a named table the Named Range is not in the Names list. 
                                // So when you create XoutputTable_ names set the range manually in the Excel name manager.
                                XOutput xout = new XOutput(nr);
                                // Add to dictionary by strip out Xinput_ prefix in key
                                dictOfOutputs.Add(xout.Name, xout);
                            }
                            break;
                        }
                    }
                }
                    if (dictOfInputs.Count==0)  
                    {
                    // Excel Author did not name any Xinput_* fields so let's try and find some possible input candidates in the set of all named fields
                    // This code should not be used other than for POCs.  Just there to make it easy for people to try random spreadseets
                    xerrorMessage = "Could not find any named input (Xinput_<i>[name]</i>) or output (Xoutput_<i>[name]</i>) fields.  Tried to find some possilbe ones from other named fields, but please consider naming cells to get desired result.";
                    foreach (IName nr in workbook.Names)
                    {
                        if (nr.RefersToRange!= null && !nr.Name.Contains("!") && nr.RefersToRange.CellCount == 1 && !nr.RefersToRange.HasFormula && dictOfInputs.Count < 10)
                        {
                            XInput xin = new XInput(nr);
                            // Add to dictionary but strip out Xinput_ prefix in key
                            dictOfInputs.Add(xin.Name, xin);
                        }
                        if (nr.RefersToRange!= null && !nr.Name.Contains("!") && nr.RefersToRange.HasFormula && nr.RefersToRange.CellCount < 100 && dictOfOutputs.Count < 20)
                        {
                            //Could be an Output that not a super large range and we have identified less than 20, so let's add - trimming down name to strip "XOutput_".
                            XOutput xout = new XOutput(nr);
                            // Add to dictionary by strip out Xinput_ prefix in key
                            dictOfOutputs.Add(xout.Name, xout);
                        }
                    }

                }

            memUsed = GC.GetTotalMemory(true) - memUsed;
            length = value.Length;

            }
        }
        public int Length
        {
            get
            {
                return length;
            }
        }
        public string Hash
        {
            get
            {
                return md5hash;
            }
        }
        public long MemUsed
        {
            get
            {
                return memUsed;
            }
        }
        public bool XdoNotLog
        {
            get
            {
                return xdoNotLog;
            }
        }
        public bool XdynamicTestUI
        {
            get
            {
                return xdynamicTestUI;
            }
        }
        public string Xauthor
        {
            get
            {
                return xauthor;
            }
        }
        public string XreleaseDate
        {
            get
            {
                return xreleaseDate;
            }
        }
        public string Xtags
        {
            get
            {
                return xtags;
            }
        }
        public string Xnotes
        {
            get
            {
                return xnotes;
            }
        }
        public string Xrevision
        {
            get
            {
                return xrevision;
            }
        }
        public string Xname
        {
            get
            {
                if (xname.Equals(""))
                {
                    return "XLE:" + Hash; 
                } else
                {
                    return xname;
                }

            }
        }
        public string XlogMessage
        {
            get
            {
                return xlogMessage;
            }
        }
        public string XerrorMessage
        {
            get
            {
                return xerrorMessage;
            }
        }

        public Dictionary<string,XInput> Xinputs
        {
            get
            {
                return dictOfInputs;
            }
        }
        public Dictionary<string, XOutput> Xoutputs
        {
            get
            {
                return dictOfOutputs;
            }
        }
        public List<XInput> lstXinputs
        {
            get
            {
                return dictOfInputs.Values.ToList();
            }
        }

        public List<XOutput> lstXoutputs
        {
            get
            {
                return dictOfOutputs.Values.ToList();
            }
        }
        public SpreadsheetGear.IWorkbook Workbook
        {
            get
            {
                return workbook;
            }
        }
        public string FileName { get; set; }

        public double LastCalctime
        {
            get
            {
                return lastCalctime;
            }

        }

        public string DebugMessage
        {
            get
            {
                return debugMessage;
            }

            set
            {
                debugMessage = value;
            }
        }

        public void LoadFromFile(string fName)
        {
            FileStream fsSource = new FileStream(fName, FileMode.Open, FileAccess.Read);
            Data = Helper.ReadFully(fsSource, 0);
            FileName = Regex.Match(fName, @"(/|\\)?(?<fileName>[^(/|\\)]+)$").Groups["fileName"].ToString();
            wasLoadedFromFile = true;
            // TODO Make sure we don't uploade the same has twice

        }

        /// <summary>
        /// Calls the calculate function on the spreasheet.    Check XInputs for excel input value validation errors and XerrorMessage for user specfied 
        /// errors supplied to XerrorMessage named field
        /// </summary>
        /// <returns>xlog object with a copy of all inputs and outputs as well as other in.
        ///  please remember to set other parameters such as tags, purpose, etc on log</returns>
        public XLog Calculate(string sSystem, string sPurpose)
        {
            System.Diagnostics.Stopwatch stopWatch = new System.Diagnostics.Stopwatch();
            stopWatch.Start();

            // We have a tight stack under some asp.net configs (256K) so need to move recusion limit down from 1024 or else you get stack overflows on very large spreadsheets. 
            // Performance still seems to be OK, but it becomes a problem, we could also create a new thread with a larger stack and move the
            // calc to it.   Google "asp.net increase stack size" for the snippet.
            workbook.WorkbookSet.MaxRecursions = 512; 

            workbook.WorkbookSet.Calculate();
            stopWatch.Stop();
            lastCalctime = (stopWatch.Elapsed.TotalMilliseconds * 1.0);

            //Update value of error message member from any defined error message file
            if (workbook.Names["XerrorMessage"]?.RefersToRange!=null) xerrorMessage = workbook.Names["XerrorMessage"].RefersToRange.Text;
 
            //Update value of log message member. 
            if (workbook.Names["XlogMessage"]?.RefersToRange != null) xlogMessage = workbook.Names["XlogMessage"].RefersToRange.Text;

            // creates a new Xlog and seeds it values with the details from this Excel objects's sate
            XLog xlogOut = new XLog(this);

            // Append an input error messages to the log error message (would presntly only contain the excel level XErrorMessag if explicitly set in the excel)
            foreach (XInput xin in lstXinputs)
            {
                if (xin.ErrorMessage.Length > 0) xlogOut.ErrorMessage += " | (Input " + xin.Name + "): " + xin.ErrorMessage;
            }

            xlogOut.CallPurpose = sPurpose;
            xlogOut.SourceSystem = sSystem;
            xlogOut.CalcTime = LastCalctime;
            
           
            return xlogOut;
        }

        public void resetToDefalutInputs ()
        {
            foreach (KeyValuePair<string, XInput> entry in dictOfInputs)
            {
                entry.Value.resetToDefaultValue();
            }
        }
 
		public string ToJSON(XLog xLog)
		{
			dynamic jsonObject = new JObject();
			jsonObject.inputs = this.lstXinputs;
			jsonObject.outputs = this.lstXoutputs;
			jsonObject.log = xLog;
			// jsonObject.graphs = xGraph;
			return jsonObject.ToString();
		}

    }
}