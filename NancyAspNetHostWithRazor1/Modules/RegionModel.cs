using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Serialization;
using System.IO;
namespace NancyAspNetHostWithRazor1.Modules
{
	public class RegionModel
	{
		public byte IpSecondByte;
		public string Passkey;
		public string Name;
		public double RegionalCoeff;
		public double EquipmentCoeff;
	}
	public class RegionsModel
	{
		public List<RegionModel> RegionModels=new List<RegionModel>();
		public void Serialize(string filename)
		{
			XmlSerializer serializer = new XmlSerializer(this.RegionModels.GetType());
			using (FileStream fs = new FileStream(filename,FileMode.Create)) {
				serializer.Serialize(fs,this.RegionModels);
			}
		}
		public void Deserialize(string filename)
		{
			XmlSerializer serializer = new XmlSerializer(this.RegionModels.GetType());
			using (FileStream fs = new FileStream(filename, FileMode.Open))
			{
				this.RegionModels=(List<RegionModel>)serializer.Deserialize(fs);
			}
		}
	}
}