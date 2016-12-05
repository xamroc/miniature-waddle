using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using SpreadsheetGear;

namespace PRoschke_ExcelTest3
{
    /// <summary>
    /// Class for holding and managing the in-memory Excel files and spreadsheet gear workbook instance for Test beds and runs.
    /// </summary>
    public class XLTest : IEquatable<XLTest>
    {
        private byte[] data;
        private string md5hash = "";
        private SpreadsheetGear.IWorkbook workbook;
        private long memUsed = 0;
        private int length = 0;

        //Values extracted from One-of named ranges at load time (ie. these should never change during usage)

        private string debugMessage = "";  //just used to held diagnose parsing errors.
        private bool xdoNotLog = false;

    

        // These Valus are dynamic and should change during usage... should be reset after earch Calculate funtion
        // TODO  move calculation execution a function of this class
        private string xerrorMessage = "";
        private string xlogMessage = "";

        private Dictionary<string, IRange> dictTestInputs = new Dictionary<string, IRange>();
        private Dictionary<string, IRange> dictTestOutputs = new Dictionary<string, IRange>();
        private Dictionary<string, IRange> dictTestExpecteds = new Dictionary<string, IRange>();

        /// <summary>
        /// Holds the testbed input Excel file
        /// </summary>
        /// <param name="b">Byte array of the Excel file that contains the inputs</param>
        public XLTest(byte[] b)
        {
            memUsed = GC.GetTotalMemory(true);
            md5hash = Helper.GetMD5Hash(b);
            data = b;
            // Create a new empty workbook set.
            SpreadsheetGear.IWorkbookSet workbookSet = SpreadsheetGear.Factory.GetWorkbookSet();
            // Open the saved workbook from memory.
            workbook = workbookSet.Workbooks.OpenFromMemory(data);

            memUsed = GC.GetTotalMemory(true) - memUsed;
            length = b.Length;

            //Get our dictionary Inputs
            foreach (IName nr in workbook.Names)
            {
                //debugMessage += "Found iName: " + nr.Name + " | ";
                if (nr.RefersToRange != null)
                { // you can delete cells but still have the refernce there.
                    if (nr.Name.StartsWith("XinputTemplate_"))
                    {
                        dictTestInputs.Add(nr.Name.Substring(15), nr.RefersToRange);
                    }
                    else if (nr.Name.StartsWith("XoutputTemplate_"))
                    {
                        dictTestOutputs.Add(nr.Name.Substring(16), nr.RefersToRange);
                    }
                    else if (nr.Name.StartsWith("XexpectedResultTemplate_"))
                    {
                        dictTestExpecteds.Add(nr.Name.Substring(24), nr.RefersToRange);
                    }
                }
            }
        }
                

        /// <summary>
        /// Loads the Test bed from an Excel file on disk  
        /// </summary>
        /// <param name="fName"></param>
        public XLTest(string fName) : this(Helper.ReadFully(new FileStream(fName, FileMode.Open, FileAccess.Read), 0))  // calls the byte [] constructor after reading file
        {
            FileName = Regex.Match(fName, @"(/|\\)?(?<fileName>[^(/|\\)]+)$").Groups["fileName"].ToString();
        }

        public override string ToString()
        {
            return "Report Hash: " + md5hash.ToString();
        }
        public override bool Equals(object obj)
        {
            return (this == obj);
        }

        public bool Equals(XLTest other)
        {
            if (other == null) return false;
            return (this.md5hash.Equals(other.md5hash));
        }

        public static bool operator !=(XLTest a, XLTest b)
        {

            return !(a == b);
        }

        public static bool operator ==(XLTest a, XLTest b)
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
                return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";  //matches the  OpenXMLWorkbook save type of spreadsheet gear
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
 


        public Dictionary<string, IRange> TestInputs
        {
            get
            {
                return dictTestInputs;
            }
        }
        public Dictionary<string, IRange> TestOutputs
        {
            get
            {
                return dictTestOutputs;
            }
        }
        public Dictionary<string, IRange> TestExpected
        {
            get
            {
                return dictTestExpecteds;
            }
        }
        public List<IRange> lstTestInputs
        {
            get
            {
                return dictTestInputs.Values.ToList();
            }
        }

        public List<IRange> lstTestOutputs
        {
            get
            {
                return dictTestOutputs.Values.ToList();
            }
        }
        public List<IRange> lstTestExpecteds
        {
            get
            {
                return dictTestExpecteds.Values.ToList();
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

        public bool XdoNotLog
        {
            get
            {
                return xdoNotLog;
            }

            set
            {
                xdoNotLog = value;
            }
        }

        public string XtestTargetService
        {
            get
            {
                return GetNamedField("XtestTargetService");
            }

            set
            {
                SetNamedField("XtestTargetService", value);
            }
        }

        public string XtestTargetRevision
        {
            get
            {
                return GetNamedField("XtestTargetRevision");
            }

            set
            {
                SetNamedField("XtestTargetRevision", value);
            }
        }

        public string XtestTargetEngine
        {
            get
            {
                return GetNamedField("XtestTargetEngine");
            }

            set
            {
                SetNamedField("XtestTargetEngine", value);
            }
        }

        public string XtestTargetURL
        {
            get
            {
                return GetNamedField("XtestTargetURL");
            }

            set
            {
                SetNamedField("XtestTargetURL", value);
            }
        }

        public string XtestTags
        {
            get
            {
                return GetNamedField("XtestTags");
            }

            set
            {
                SetNamedField("XtestTags", value);
            }
        }

        public string XtestPurpose
        {
            get
            {
                return GetNamedField("XtestPurpose");
            }

            set
            {
                SetNamedField("XtestPurpose", value);
            }
        }

        public string XtestID
        {
            get
            {
                return GetNamedField("XtestID");
            }

            set
            {
                SetNamedField("XtestID", value);
            }
        }

        public string XtestedService
        {
            get
            {
                return GetNamedField("XtestedService");
            }

            set
            {
                SetNamedField("XtestedService", value);
            }
        }

        public string XtestedRevision
        {
            get
            {
                return GetNamedField("XtestedRevision");
            }

            set
            {
                SetNamedField("XtestedRevision", value);
            }
        }

        public string XtestedEngine
        {
            get
            {
                return GetNamedField("XtestedEngine");
            }

            set
            {
                SetNamedField("XtestedEngine", value);
            }
        }

        public string XtestedURL
        {
            get
            {
                return GetNamedField("XtestedURL");
            }

            set
            {
                SetNamedField("XtestedURL", value);
            }
        }



        public int XtestsRun
        {
            get
            {
                return  workbook.Names["XtestsRun"]?.RefersToRange?.Value != null ? (int)workbook.Names["XtestsRun"].RefersToRange.Value : 0;
            }

            set
            {
                if (workbook.Names["XtestsRun"]?.RefersToRange != null)
                {
                    workbook.Names["XtestsRun"].RefersToRange.Value = value;
                }
            }
        }

        public int XtestErrors
        {
            get
            {
                return workbook.Names["XtestErrors"]?.RefersToRange?.Value != null ? (int)workbook.Names["XtestErrorsn"].RefersToRange.Value : 0;
            }

            set
            {
                if (workbook.Names["XtestErrors"]?.RefersToRange != null)
                {
                    workbook.Names["XtestErrors"].RefersToRange.Value = value;
                }
            }
        }

        public double XtestMatch
        {
            get
            {
                return workbook.Names["XtestMatch"]?.RefersToRange?.Value != null ? (double)workbook.Names["XtestMatch"].RefersToRange.Value : 0;
            }

            set
            {
                if (workbook.Names["XtestMatch"]?.RefersToRange != null)
                {
                    workbook.Names["XtestMatch"].RefersToRange.Value = value;
                }
            }
        }

        public int XtestExecutionTime
        {
            get
            {
                return workbook.Names["XtestExecutionTime"]?.RefersToRange != null ? (int)workbook.Names["XtestExecutionTime"].RefersToRange.Value : 0;
            }

            set
            {
                if (workbook.Names["XtestExecutionTime"]?.RefersToRange != null)
                {
                    workbook.Names["XtestExecutionTime"].RefersToRange.Value = value;
                }
            }
        }

        public IRange XtestReferenceTemplate
        {
            get
            {
                return workbook.Names["XtestReferenceTemplate"]?.RefersToRange;
            }

        }

        public IRange XtestInputTemplate
        {
            get
            {
                return workbook.Names["XtestInputTemplate"]?.RefersToRange;
            }

         }

        public IRange XtestOutputTemplate
        {
            get
            {
                return workbook.Names["XtestOutputTemplate"]?.RefersToRange;
            }

        }

        public IRange XtestExpectedResultTemplate
        {
            get
            {
                return workbook.Names["XtestExpectedResultTemplate"]?.RefersToRange;
            }  
        }

        public IRange XtestLogDetailsTemplate
        {
            get
            {
                return workbook.Names["XtestLogDetailsTemplate"]?.RefersToRange;
            }

        }
        public IRange XtestErrorDetailsTemplate
        {
            get
            {
                return workbook.Names["XtestErrorDetailsTemplate"]?.RefersToRange;
            }

        }


        /// <summary>
        /// Calls the calculate function on the spreasheet.    Check XInputs for excel input value validation errors and XerrorMessage for user specfied 
        /// errors supplied to XerrorMessage named field
        /// </summary>
        /// <param name="sSystem">CAlling system indentifier to add to the log entries generated</param>
        /// <param name="sPurpose">Purpose of this test to add to the log entries</param>
        /// <param name="XLObjs">List of in memor XL objects to try and find the source matching calc engine in</param>
        /// <param name="Xlogs">List of logs to append test results to</param>
        public void ExecuteTests(string sSystem, string sPurpose, List<XLObj> XLObjs, List<XLog> Xlogs)
        {
            // Let's see if we can find the referred to engine in the test set excel

            XLObj XLO = XLObjs[0];
            //if (XLO == null)
            //    XtestTargetEngine = "DID NOT FIND " + XtestTargetEngine;
            //{
            //    XLO = XLObjs.Find(x => (x.Xname == XtestTargetService && x.Xrevision == XtestTargetRevision));
            //}
            //if (XLO == null)
            //{
            //    XLO = XLObjs.Find(x => (x.Xname == XtestTargetService));  // last try-- see if we can at least find a matching service.
            //}
            //if (XLO == null) throw new System.Exception("Can't find test target engine.");

            XtestPurpose = sPurpose;

            System.Diagnostics.Stopwatch stopWatch = new System.Diagnostics.Stopwatch();
            stopWatch.Start();

            lock (XLO)  // Internal excel model is stateful - so we are creating a critical section here to prevent threading problems.
            {
                // Iterate over test cases in the file
                bool bDataInRow = true;
                int i = 0;
                int MatchErrors = 0;
                int Errors = 0;
                bool bRowInputValudationErrors = false;
                bool bRowMatchErrors = false;
                while (bDataInRow == true)
                {
                    bDataInRow = false;
                    bRowInputValudationErrors = false;
                    bRowMatchErrors = false;
                    XLO.resetToDefalutInputs();
                    foreach (KeyValuePair<string, IRange> entry in dictTestInputs)
                    {
                        if (entry.Value.Offset(i, 0).Value?.ToString() == null || entry.Value.Offset(i, 0).Value.ToString().Equals(""))
                        {
                            // The test excel did not specify a value... get the default from the engine
                            entry.Value.Offset(i, 0).Value = XLO.Xinputs[entry.Key].Value;
                            entry.Value.Offset(i, 0).Font.Color = SpreadsheetGear.Colors.Gray;
                        }
                        else
                        {
                            XLO.Xinputs[entry.Key].Value = entry.Value.Offset(i, 0).Value.ToString();
                            if (!XLO.Xinputs[entry.Key].IsValid)
                            {
                                //We have some input errors
                                bRowInputValudationErrors = true;
                                entry.Value.Offset(i, 0).Font.Color = SpreadsheetGear.Colors.Red;
                            }
                            bDataInRow = true;   // we found some input values in the curent row.
                        }
                    }
                    XLog xl;
                    xl = XLO.Calculate("Excel Engine UI Test Harness", "ManualTest");
                    xl.CallPurpose = sPurpose;
                    xl.SourceSystem = sSystem;
                    Xlogs.Add(xl);

                    // Truncate Log list if necessary (drop entries from front so most recent ones are kept)
                    if (Xlogs.Count > 10000) Xlogs.RemoveRange(0, 1000);  // given the internal structure of C# lists... this is actually pretty efficient.


                    foreach (XOutput xo in XLO.lstXoutputs)
                    {
                        if (dictTestOutputs[xo.Name] != null)
                        {
                            // We have a place to put the output
                            dictTestOutputs[xo.Name].Offset(i, 0).Value = xo.Value;
                            if (dictTestExpecteds[xo.Name]?.Offset(i, 0).Value?.ToString() != null && !dictTestExpecteds[xo.Name].Offset(i, 0).Value.ToString().Equals(""))
                            {
                                // the spreadsheet also has an expeted result entry
                                if (!dictTestExpecteds[xo.Name].Offset(i, 0).Text.Equals(xo.Value.Trim()))
                                {
                                    // ... but the value doesn't match
                                    dictTestExpecteds[xo.Name].Offset(i, 0).Interior.Color = SpreadsheetGear.Colors.Red;
                                    bRowMatchErrors = true;

                                }
                            }
                        }

                    }
                    if (XtestErrorDetailsTemplate != null)
                    {
                        XtestErrorDetailsTemplate.Offset(i, 0).Value = xl.ErrorMessage;
                        XtestErrorDetailsTemplate.Offset(i, 0).Font.Color = SpreadsheetGear.Colors.Red;
                        if (!xl.ErrorMessage.Equals("") || bRowInputValudationErrors) Errors++;
                        if (bRowMatchErrors) MatchErrors++;
                    }
                    if (XtestLogDetailsTemplate != null)
                    {
                        XtestLogDetailsTemplate.Offset(i, 0).Value = xl.LogMessage;
                    }
                    if (XtestReferenceTemplate != null) XtestReferenceTemplate.Offset(i, 0).Value = "ID:" + (i + 1);

                    i++;

                }
                                 
            stopWatch.Stop();
            XtestExecutionTime = (int)stopWatch.Elapsed.TotalMilliseconds;
            XtestsRun = i;
            XtestErrors = Errors;
            XtestMatch = 1.0 * (i - MatchErrors-1) / i;
            XtestedEngine = XLO.Hash;
            XtestedRevision = XLO.Xrevision;
            XtestedService = XLO.Xname;
            UpdateMemoryFile();
            }  // End Critical section

            return;
        }
 
        /// <summary>
        /// Modifies the in memory test template to match a particular engine.   Should be called on objects just loaded with the blank Text set template.
        /// </summary>
        /// <param name="XLO">The engine to match</param>
        /// <returns></returns>
        public XLTest TemplateToOutput (XLObj XLO)
        {
            XLTest NXT = new XLTest(this.data);


            // Expand Input portion of text case table to inlcude a list of all inputs for XLO
            List<XInput> Xis = XLO.lstXinputs;
            Xis.Reverse();   // ensures names come out in alphabeticl order
            foreach (XInput xi in Xis)
            {
                NXT.XtestInputTemplate.Offset(-1,1).Cells.Insert(InsertShiftDirection.Right);   // copies formatting
                NXT.XtestInputTemplate.Offset(0,1).Cells.Insert(InsertShiftDirection.Right);   // also copies formatting
                NXT.XtestInputTemplate.Offset(-1, 1).Value = xi.Name;
                NXT.Workbook.Names.Add("XinputTemplate_" + xi.Name, "="+NXT.XtestInputTemplate.Offset(0, 1).Address);   
            }
            NXT.XtestInputTemplate.Offset(-1, 0).Delete(DeleteShiftDirection.Left);
            NXT.XtestInputTemplate.Delete(DeleteShiftDirection.Left);

            // Do the same for Xoutputs.
            List<XOutput> Xos = XLO.lstXoutputs;
            Xos.Reverse();   // ensures names come out in alphabeticl order
            foreach (XOutput xo in Xos)
            {
                NXT.XtestOutputTemplate.Offset(-1, 1).Cells.Insert(InsertShiftDirection.Right);
                NXT.XtestOutputTemplate.Offset(0, 1).Cells.Insert(InsertShiftDirection.Right);
                NXT.XtestOutputTemplate.Offset(-1, 1).Value = xo.Name;
                NXT.Workbook.Names.Add("XoutputTemplate_" + xo.Name, "=" + NXT.XtestOutputTemplate.Offset(0, 1).Address);
            }
            NXT.XtestOutputTemplate.Offset(-1, 0).Delete(DeleteShiftDirection.Left);
            NXT.XtestOutputTemplate.Delete(DeleteShiftDirection.Left);

            // Do the same for expectedresults
            foreach (XOutput xo in Xos)
            {
                NXT.XtestExpectedResultTemplate.Offset(-1, 1).Cells.Insert(InsertShiftDirection.Right);
                NXT.XtestExpectedResultTemplate.Offset(0, 1).Cells.Insert(InsertShiftDirection.Right);
                NXT.XtestExpectedResultTemplate.Offset(-1, 1).Value = xo.Name;
                NXT.Workbook.Names.Add("XexpectedResultTemplate_" + xo.Name, "=" + NXT.XtestExpectedResultTemplate.Offset(0, 1).Address);
            }
            NXT.XtestExpectedResultTemplate.Offset(-1, 0).Delete(DeleteShiftDirection.Left);
            NXT.XtestExpectedResultTemplate.Delete(DeleteShiftDirection.Left);

            //Set other parmeters
            NXT.XtestTargetEngine = XLO.Hash;
            NXT.XtestTargetRevision = XLO.Xrevision;
            NXT.XtestTargetService = XLO.Xname;
            
            //Set an appropirate name for later downlaod.
            NXT.FileName = NXT.XtestTargetService + "-TestSetTemplate.xlsx";

            NXT.UpdateMemoryFile();
            return NXT;
        }

        public void UpdateMemoryFile()
        {

            // We have a tight stack under some asp.net configs (256K) so need to move recusion limit down from 1024 or else you get stack overflows on very large spreadsheets. 
            // Performance still seems to be OK, but it becomes a problem, we could also create a new thread with a larger stack and move the
            // calc to it.   Google "asp.net increase stack size" for the snippet.
            workbook.WorkbookSet.MaxRecursions = 512;

            workbook.WorkbookSet.Calculate();
            data = workbook.SaveToMemory(FileFormat.OpenXMLWorkbook);
        }

        public void LoadTestCasesFromLogs(List<XLog> xlogs)
        {
            int i = 0;
            foreach (XLog xl in xlogs)
            {

                lstTestInputs[0].Offset(i+1,0).EntireRow.Insert(InsertShiftDirection.Down); //make a new row under template row

                foreach (KeyValuePair<string, IRange> xi in dictTestInputs)
                {
                    string sTry = "";
                    xl.Xinputs.TryGetValue(xi.Key, out sTry);
                    xi.Value.Offset(i, 0).Formula = sTry==null?"":sTry;
                }
                foreach (KeyValuePair<string, IRange> xo in dictTestExpecteds)
                {
                    string sTry = "";
                    xl.Xoutputs.TryGetValue(xo.Key, out sTry);
                    xo.Value.Offset(i, 0).Formula = sTry==null?"":sTry;
                }
                XtestReferenceTemplate.Offset(i, 0).Formula = xl.TimeStamp;
                XtestErrorDetailsTemplate.Offset(i, 0).Formula = xl.ErrorMessage;
                XtestLogDetailsTemplate.Offset(i, 0).Formula = xl.LogMessage;
                i++;
            }
        }

        /// <summary>
        /// Returns the string value of the field or "" if not found
        /// </summary>
        /// <param name="fn">Name of Named field in excel file</param>
        /// <returns></returns>
        private string GetNamedField (string fn)
        {
            return workbook.Names[fn]?.RefersToRange?.Value != null ? workbook.Names[fn].RefersToRange.Value.ToString() : ""; 
        }

        /// <summary>
        /// Stets the value of a naned cell in Excel.  
        /// </summary>
        /// <param name="fn">Name of Named field in excel file</param>
        /// <param name="val">value to set it to</param>
        /// <returns>False if name not found</returns>
        private bool SetNamedField(string fn, string val)
        {
            if (workbook.Names[fn]?.RefersToRange != null)
            {
                workbook.Names[fn].RefersToRange.Value = val;
                return true;
            } else
            {
                return false;
            }
        }
    }
}