using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NodaTime.TimeZones;

namespace NaumenPenalty.Model
{
    public class NaumenTimeZone
    {
        public long Id { get; set; }
        public string Code { get; set; }
        public TimeZoneInfo TimeZoneInfo
        {
            get
            {
                var tzMappings = TzdbDateTimeZoneSource.Default.WindowsMapping.MapZones;
                var map = tzMappings.FirstOrDefault(x => x.TzdbIds.Any(y => y.Equals(Code, StringComparison.OrdinalIgnoreCase)));
                return TimeZoneInfo.FindSystemTimeZoneById(map.WindowsId);
            }
        }
    }
}
