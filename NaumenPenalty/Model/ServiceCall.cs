using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Dapper;
namespace NaumenPenalty.Model
{
	public class ServiceCall
	{
		public long Id { get; set; }
		public bool Removed { get; set; }
		public DateTime RegistrationDate { get; set; }
		public DateTime DeadLineTime { get; set; }
		public bool MassProblem { get; set; }
		public long Number { get; set; }
		public long ResolutionTime { get; set; }
		public String State { get; set; }
		public long ClientouId { get; set; }
		public long TimeZoneId { get; set; }
		public long AgreementId { get; set; }
		public Agreement Agreement { get; set; }
		public Ou ClientOu{ get; set; }
		public NaumenTimeZone NaumenTimeZone{ get; set; }
        public DateTime StateStartTime { get; set; }
        public long ServiceId { get; set; }
	}
}