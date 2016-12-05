using System;
using System.IO;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using SpreadsheetGear;
using Mono.Options;
using Newtonsoft.Json;
using System.Linq;
using Newtonsoft.Json.Linq;
using NetMQ;
using NetMQ.Sockets;

namespace PRoschke_ExcelTest3
{
    class Program
    {
        private static void Main(string[] args)
        {
            String inputFilename = "";
            String bulkTestFilename = "";
			String jsonString = "";
            bool bulkTesting = false;
			bool jsonOverride = false;
			bool testTemplate = false;
			bool fieldsGeneration = false;

			List<XLObj> XLObjs = new List<XLObj>();
			List<XLog> XLogs = new List<XLog>();
			List<XInput> XInputs = null;

            Dictionary<string, XInput> dictOfInputs = new Dictionary<string, XInput>();
            Dictionary<string, XOutput> dictOfOutputs = new Dictionary<string, XOutput>();

            var show_help = false;
            var os = new OptionSet () {
                { "i|input=", "the input XLS filename",
                   v => inputFilename = v },
                { "b|bulkrun=", "the bulkrun test sheet XLS filename",
                  v => {
                    bulkTestFilename = v;
                    bulkTesting = true;
                  }
                },
				{ "j|json=", "the input-override json string",
				  v => {
					jsonString = v;
					jsonOverride = true;
				  }
				},
				{ "t|testtemplate", "test template generation",
				  v => {
					testTemplate = true;
				  }
				},
				{ "f|fields", "fields generation",
				   v => {
					fieldsGeneration = true;
				  }
				},
                { "h|help",  "show this message and exit", 
                   v => show_help = v != null },
            };

            List<string> extra;
			try
			{
				extra = os.Parse(args);
			}
			catch (OptionException e)
			{
				Console.Error.Write("Error: ");
				Console.WriteLine(e.Message);
				Console.WriteLine("Try `hsece-cli --help' for more information.");
				return;
			}

			try
			{
				bool isPricer = false;

				FileStream ifs = new FileStream(inputFilename, FileMode.Open);
				XLObj XLO = new XLObj()
				{
					FileName = Regex.Match(inputFilename, @"(/|\\)?(?<fileName>[^(/|\\)]+)$").Groups["fileName"].ToString(),   // to trim off whole path from browsers like IE
					MimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
					Data = Helper.ReadFully(ifs, 0)
				};
				XLog xLog = null;

				// check to see if same engine signature already in memory, otherwise add to list
				if (XLObjs.Find(x => x.Hash == XLO.Hash) == null) XLObjs.Add(XLO);

				using (var serverSock = new ResponseSocket())
				{
					serverSock.Bind("tcp://127.0.0.1:5556");

					while (true)
					{
						string fromClientMessage = serverSock.ReceiveString();

						// write in stuff from JSON
						dynamic d = JObject.Parse(fromClientMessage);

						if (d != null)
						{	
							dynamic inputs = d.input.data;

							if (inputs == null)
							{
								Exception e = new Exception("No inputs override found for JSON override");
								throw e;
							}

							//Parse Inputs 
							foreach (dynamic i in inputs)
							{
								dynamic dataType = i.DataType;
								dynamic name = i.Name;
								dynamic value = i.Value;
								dynamic valueType = i.ValueType;
								dynamic htmlDataType = i.HTMLDataType;

								XInput xiSpreadsheet = XLO.Xinputs[(string)name];

								// Write it to the Excel model
								if (xiSpreadsheet.HTMLDataType.Equals("checkbox"))
								{
									//checkbox is a pain because uncheked ones don't post.  Not a beautiful soluiont below but it works.
									xiSpreadsheet.Value = value.Equals("0,1") ? "TRUE" : "FALSE";
								}
								else
								{
									xiSpreadsheet.Value = value;
								}
							}
						}

						xLog = XLO.Calculate(inputFilename, "Test Purpose");
						var processResult = new ProcessResult(xLog, XLO);
						serverSock.Send(processResult.toJSONString());
					}
				}
			}
			catch (Exception e)
			{
				var processException = new ProcessError(e);
				Console.Error.WriteLine(processException.toJSONString());
				Environment.Exit(1);
			}
        }
    }
}
