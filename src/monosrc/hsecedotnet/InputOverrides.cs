using System;
using Newtonsoft.Json;
using System.Data;
using System.Collections.Generic;

namespace PRoschke_ExcelTest3
{
	public class InputOverrides
	{
    	[JsonProperty(PropertyName = "input")]
		public List<XInput> input { get; set; }
	}
}
