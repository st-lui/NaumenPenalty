using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NaumenPenalty.Model
{
	public class Agreement
	{
		public long Id { get; set; }
		public string Title { get; set; }
		public IEnumerable<ServiceCall> ServiceCalls { get; set; }
	}
}
