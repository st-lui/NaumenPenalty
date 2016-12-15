using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NaumenPenalty.Model
{
	public class Ou
	{
		public long Id { get; set; }
		public long LocationId { get; set; }
		public Location Location { get; set; }
		public IEnumerable<ServiceCall> ServiceCalls { get; set; }
	}
}
