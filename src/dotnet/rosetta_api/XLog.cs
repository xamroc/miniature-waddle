using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace PRoschke_ExcelTest3
{
    public class XLog
    {
        private string engineID ="";
        private string serviceName ="";
        private string serviceRevision = "";
        private DateTime timestamp;
        private Dictionary<string, string> dictXoutputs = new Dictionary<string, string>();
        private Dictionary<string, string> dictXinputs = new Dictionary<string, string>();
        private string sourceSystem = "";
        private string callPurpose ="NA"; // RegressionTest, ManualTest, Production, QATest, UATTest, NA
        private string tags = "";
        private string logMessage = "";
        private string errorMessage ="";
        private double calcTime = 0.0;

        public XLog(XLObj XLO)
        {
            engineID = XLO.Hash;
            serviceName = XLO.Xname;
            ServiceRevision = XLO.Xrevision;
            timestamp = DateTime.Now;
            foreach (var xik in XLO.Xinputs.Keys)
            {
                dictXinputs.Add(xik, XLO.Xinputs[xik].Value);
            }
            timestamp = DateTime.Now;
            foreach (var xok in XLO.Xoutputs.Keys)
            {
                dictXoutputs.Add(xok, XLO.Xoutputs[xok].Value);
            }
            logMessage = XLO.XlogMessage;
            errorMessage = XLO.XerrorMessage;
            
        }

        public XLog ()
        {
            timestamp = DateTime.Now;
        }

        public XLog (Dictionary<string, XInput> xin, Dictionary<string, XOutput> xout)
        {
            timestamp = DateTime.Now;
            foreach (var xik in xin.Keys)
            {
                dictXinputs.Add(xik, xout[xik].Value);
            }
            timestamp = DateTime.Now;
            foreach (var xok in xout.Keys)
            {
                dictXoutputs.Add(xok, xout[xok].Value);
            }
        }

        public string EngineID
        {
            get
            {
                return engineID;
            }

            set
            {
                engineID = value;
            }
        }

        public string ServiceName
        {
            get
            {
                return serviceName;
            }

            set
            {
                serviceName = value;
            }
        }

        public string ServiceRevision
        {
            get
            {
                return serviceRevision;
            }

            set
            {
                serviceRevision = value;
            }
        }

        public string TimeStamp
        {
            get
            {
                return timestamp.ToString();
            }

        }

        public Dictionary<string, string> Xinputs
        {
            get
            {
                return dictXinputs;
            }

            set
            {
                dictXinputs = value;
            }
        }

        public Dictionary<string, string> Xoutputs
        {
            get
            {
                return dictXoutputs;
            }

            set
            {
                dictXoutputs= value;
            }
        }

        public string SourceSystem
        {
            get
            {
                return sourceSystem;
            }

            set
            {
                sourceSystem = value;
            }
        }

        public string Tags
        {
            get
            {
                return tags;
            }

            set
            {
                tags = value;
            }
        }

        public string LogMessage
        {
            get
            {
                return logMessage;
            }

            set
            {
                logMessage = value;
            }
        }

        public string ErrorMessage
        {
            get
            {
                return errorMessage;
            }

            set
            {
                errorMessage = value;
            }
        }

        public double CalcTime
        {
            get
            {
                return calcTime;
            }

            set
            {
                calcTime = value;
            }
        }

        public string CallPurpose
        {
            get
            {
                return callPurpose;
            }

            set
            {
                callPurpose = value;
            }
        }

        public Dictionary<string, string> DictXoutputs
        {
            get
            {
                return dictXoutputs;
            }

            set
            {
                dictXoutputs = value;
            }
        }
    }
}