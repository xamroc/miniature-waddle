using System;
using System.Collections.Generic;
using System.Dynamic;
using Newtonsoft.Json;

namespace PRoschke_ExcelTest3
{
	public class ProcessError
	{
		private dynamic dynResult = new ExpandoObject();

		public ProcessError(Exception e)
		{
			dynResult.error = e.Message;
			dynResult.stack = e.StackTrace;
		}

		public String toJSONString()
		{
			return JsonConvert.SerializeObject(dynResult);
		}
	}
}
