using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NaumenPenalty.Configuration
{
	public class FeeConfigSection : ConfigurationSection
	{
		[ConfigurationProperty("Fees")]
		public FeeCollection Items => (FeeCollection)base["Fees"];
	}

	[ConfigurationCollection(typeof(FeeElement))]
	public class FeeCollection : ConfigurationElementCollection
	{
		protected override ConfigurationElement CreateNewElement()
		{
			return new FeeElement();
		}

		protected override object GetElementKey(ConfigurationElement element)
		{
			return ((FeeElement)element).Id;
		}
		public object this[object idx] => BaseGet(idx);
	}

	public class FeeElement : ConfigurationElement
	{
		[ConfigurationProperty("Id", DefaultValue = 0, IsKey = true, IsRequired = true)]
		public int Id => (int)base["Id"];
		[ConfigurationProperty("Description", DefaultValue = "", IsKey = false, IsRequired = true)]
		public string Description => (string)base["Description"];
		[ConfigurationProperty("FeeValue", DefaultValue = 0.0, IsKey = false, IsRequired = true)]
		public double FeeValue => (double)base["FeeValue"];
	}
}
