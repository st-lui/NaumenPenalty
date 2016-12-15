using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NaumenPenalty.Model
{
	public class Location
	{
		public long Id { get; set; }
		public string PostCode { get; set; }
		public string Title { get; set; }
		public long? PostalClass { get; set; }
		public string Type { get; set; }
		public long? Parent { get; set; }
		public string ShortName { get; set; }
		IEnumerable<Ou> Ous { get; set; }
	}
}
