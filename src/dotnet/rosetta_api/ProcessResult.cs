using System;
using System.Collections.Generic;
using System.Dynamic;
using Newtonsoft.Json;
namespace PRoschke_ExcelTest3
{

	public class ProcessResult
	{
		private XLObj xlObject ;
		private dynamic dynResult = new ExpandoObject();

		public ProcessResult(XLog xLog, XLObj xlObject)
		{
			dynResult.log = xLog;
			this.xlObject = xlObject;
			dynResult.input = xlObject.lstXinputs;
			dynResult.output = xlObject.lstXoutputs;

			// load all excel charts into base64 string array
			var charts = new List<ExpandoObject>();
			charts = getChartDataList();
			dynResult.charts = charts;
		}

		public String toJSONString()
		{
			return JsonConvert.SerializeObject(dynResult);
		}

		public List<ExpandoObject> getChartDataList() {
			var chartStringList = new List<ExpandoObject>();

			SpreadsheetGear.Drawing.Image imgToRender = null;
			// check all worksheets
			foreach (SpreadsheetGear.IWorksheet ws in xlObject.Workbook.Worksheets)
			{
				//for shapes
				foreach (SpreadsheetGear.Shapes.IShape shape in ws.Shapes)
				{
					if (shape.HasChart)
					{
						imgToRender = new SpreadsheetGear.Drawing.Image(shape);
					//Found a chart
#pragma warning disable XS0001 // Find usages of mono todo items
						System.IO.MemoryStream strm = new System.IO.MemoryStream();
#pragma warning restore XS0001 // Find usages of mono todo items
						imgToRender.GetBitmap().Save(strm, System.Drawing.Imaging.ImageFormat.Png);
						string base64String = Convert.ToBase64String(strm.ToArray());
						dynamic chartObject = new ExpandoObject();
						chartObject.data = base64String;
						chartStringList.Add(chartObject);
					}
				}
			}

			return chartStringList;
		}
	}
}
