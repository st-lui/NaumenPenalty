using Microsoft.VisualStudio.TestTools.UnitTesting;
using NancyAspNetHostWithRazor1.Modules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NancyAspNetHostWithRazor1.Modules.Tests
{
	[TestClass()]
	public class RegionsModelTests
	{
		[TestMethod()]
		public void SerializeTest()
		{
			RegionsModel regionsModel = new RegionsModel();
			regionsModel.RegionModels.Add(new RegionModel() { Name = "Алтайский край", Passkey = "166470ff-beb7-48a3-b2a8-540d1b7c26dc", IpSecondByte = 56, EquipmentCoeff = 0.9904, RegionalCoeff = 0.97 });
			regionsModel.Serialize("regions.xml");
			RegionsModel regionsModel1 = new RegionsModel();
			regionsModel1.Deserialize("regions.xml");
			Assert.AreEqual(regionsModel.RegionModels[0].EquipmentCoeff, regionsModel1.RegionModels[0].EquipmentCoeff);
			Assert.AreEqual(regionsModel.RegionModels[0].IpSecondByte, regionsModel1.RegionModels[0].IpSecondByte);
			Assert.AreEqual(regionsModel.RegionModels[0].Name, regionsModel1.RegionModels[0].Name);
			Assert.AreEqual(regionsModel.RegionModels[0].Passkey, regionsModel1.RegionModels[0].Passkey);
			Assert.AreEqual(regionsModel.RegionModels[0].RegionalCoeff, regionsModel1.RegionModels[0].RegionalCoeff);
		}
	}
}