using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NaumenPenalty.Configuration
{
	public class PenaltyMultiplierConfigSection : ConfigurationSection
	{
		[ConfigurationProperty("Multipliers")]
		public PenaltyMultiplierCollection Items => (PenaltyMultiplierCollection)base["Multipliers"];
	}

	[ConfigurationCollection(typeof(PenaltyMultiplierElement))]
	public class PenaltyMultiplierCollection : ConfigurationElementCollection
	{
		protected override ConfigurationElement CreateNewElement()
		{
			return new PenaltyMultiplierElement();
		}

		protected override object GetElementKey(ConfigurationElement element)
		{
			return ((PenaltyMultiplierElement)element).Index;
		}
		public object this[object idx] => BaseGet(idx);
	}

	public class PenaltyMultiplierElement : ConfigurationElement
	{
		[ConfigurationProperty("Index", DefaultValue = 0, IsKey = true, IsRequired = true)]
		public int Index
		{
			get { return (int)base["Index"]; }
			set { base["Index"] = value.ToString(); }
		}
		[ConfigurationProperty("EquipmentMultiplier", DefaultValue = 0.0, IsKey = false, IsRequired = true)]
		public double EquipmentMultiplier
		{
			get { return (double)base["EquipmentMultiplier"]; }
			set { base["EquipmentMultiplier"] = value.ToString(CultureInfo.InvariantCulture); }
		}
		[ConfigurationProperty("RegionalMultiplier", DefaultValue = 0.0, IsKey = false, IsRequired = true)]
		public double RegionalMultiplier
		{
			get { return (double)base["RegionalMultiplier"]; }
			set { base["RegionalMultiplier"] = value.ToString(CultureInfo.InvariantCulture); }
		}
	}
}
