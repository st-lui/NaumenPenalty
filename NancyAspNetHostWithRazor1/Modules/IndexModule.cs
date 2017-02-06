using System;
using Nancy;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using Nancy.Json;
using Nancy.Responses;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
namespace NancyAspNetHostWithRazor1.Modules
{
	public class IndexModule : NancyModule
	{
		string fileId;
		string resultText;
		Dictionary<byte, Tuple<string, string>> passkeys = new Dictionary<byte, Tuple<string, string>> {
			{ 56, new Tuple<string, string>("166470ff-beb7-48a3-b2a8-540d1b7c26dc", "Алтайский край") },
			{ 67, new Tuple<string, string>("b2e9acf7-09cb-4e72-8d66-588cbcfd8a30", "Республика Хакасия") } };

		static Tuple<bool, string> DownloadString(string url)
		{
			try
			{
				using (WebClient webClient = new WebClient())
				{
					var data = webClient.DownloadString(url);
					data = data.Replace("\"", @"""");
					return new Tuple<bool, string>(true, data);
				}
			}
			catch (Exception e)
			{
				MyLogger.GetInstance().Error(url + ":" + e.ToString());
				return new Tuple<bool, string>(false, url + ":" + e.ToString());
			}
		}

		static byte[] DownloadData(string url)
		{
			try
			{
				using (WebClient webClient = new WebClient())
				{
					var data = webClient.DownloadData(url);
					return data;
				}
			}
			catch (Exception e)
			{
				MyLogger.GetInstance().Error(url + ":" + e.ToString());
			}
			return null;
		}

		Tuple<bool, string> CheckReportIsReady(string url)
		{
			var reportCheck = DownloadString(url);
			if (reportCheck.Item1 == false)
			{
				return new Tuple<bool, string>(false, "Ошибка. Не удалось запросить формирование отчета.");
			}
			JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
			var checkJSON = (dynamic)javaScriptSerializer.DeserializeObject(reportCheck.Item2);
			object[] files = null;
			try
			{
				files = checkJSON["Files"];
			}
			catch (Exception ex)
			{
				MyLogger.GetInstance().Error("Ошибка парсинга ответа сервера Naumen на запрос файла отчета. В ответе не найдено поле files");
				return new Tuple<bool, string>(false, "Ошибка. Не удалось запросить файл отчета.");
			}
			if (files.Length == 0)
				return new Tuple<bool, string>(false, null);
			fileId = null;
			try
			{
				fileId = ((dynamic)files[0])["UUID"];
			}
			catch (Exception e)
			{
				MyLogger.GetInstance().Error(e.ToString());
				return new Tuple<bool, string>(false, null);
			}
			if (string.IsNullOrEmpty(fileId))
				return new Tuple<bool, string>(false, null);
			fileId = fileId.Split('$')[1];
			return new Tuple<bool, string>(true, fileId);
		}

		public string ModifyNaumenReport(string fileName, PenaltyModel penaltyModel)
		{
			if (!File.Exists(fileName))
			{
				return null;
			}
			using (FileStream fileStream = new FileStream(fileName, FileMode.Open))
			{
				IWorkbook workbook = new XSSFWorkbook(fileStream);
				ISheet activeSheet = workbook.GetSheetAt(0);
				activeSheet.ShiftRows(activeSheet.FirstRowNum, activeSheet.LastRowNum, 5);
				var currentRow = activeSheet.CreateRow(0);
				currentRow.CreateCell(1).SetCellValue("Региональный коэффициент");
				currentRow.CreateCell(2).SetCellValue(penaltyModel.RegionCoeff);
				currentRow = activeSheet.CreateRow(1);
				currentRow.CreateCell(1).SetCellValue("Коэффициент оснащенности");
				currentRow.CreateCell(2).SetCellValue(penaltyModel.EquipCoeff);
				currentRow = activeSheet.CreateRow(3);
				currentRow.CreateCell(1).SetCellValue("Сумма штрафов");

				currentRow = activeSheet.GetRow(5);
				currentRow.CopyCell(11, 12).SetCellValue("Базовая абонплата");
				currentRow.CopyCell(11, 13).SetCellValue("Размер штрафа");
				currentRow = activeSheet.GetRow(6);
				currentRow.CopyCell(11, 12).SetCellValue("13");
				currentRow.CopyCell(11, 13).SetCellValue("14");
				activeSheet.SetColumnWidth(12, activeSheet.GetColumnWidth(11));
				activeSheet.SetColumnWidth(13, activeSheet.GetColumnWidth(11));
				

				string path = Path.GetDirectoryName(fileName), fnwe = Path.GetFileNameWithoutExtension(fileName), ext = Path.GetExtension(fileName);
				string modeFilename = $"{fnwe}_mod{ext}";
				int dataStartRow = 7;
				int startColumnIndex = 0;
				int nameColumnIndex = 1;
				int qualityColumnIndex = 9;
				int coeffColumnIndex = 10;
				int classColumnIndex = 11;
				int priceColumnIndex = 12;
				int penaltyColumnIndex = 13;
				var doubleFormat = workbook.CreateDataFormat().GetFormat("##0.00");
				var cs = workbook.CreateCellStyle();
				cs.CloneStyleFrom(activeSheet.GetRow(7).Cells[coeffColumnIndex].CellStyle);
				cs.DataFormat = doubleFormat;
				currentRow = activeSheet.GetRow(dataStartRow);
				int i = 0;
				while (currentRow.GetCell(0).ToString() != "")
				{
					string opsClass = currentRow.GetCell(11).ToString();
					string abonCell = CellReference.ConvertNumToColString(priceColumnIndex) + (currentRow.RowNum + 1);
					string penaltyCoefCell = CellReference.ConvertNumToColString(coeffColumnIndex) + (currentRow.RowNum + 1);
					string qualityCell = CellReference.ConvertNumToColString(qualityColumnIndex) + (currentRow.RowNum + 1);

					string opsName = currentRow.GetCell(1).ToString();
					string price = GetPriceFromNameAndClass(nameOps: currentRow.Cells[nameColumnIndex].ToString(), opsClass: currentRow.Cells[classColumnIndex].ToString());

					currentRow.CopyCell(classColumnIndex, priceColumnIndex).SetCellValue(price);
					var cell = currentRow.CopyCell(coeffColumnIndex, penaltyColumnIndex);
					cell.SetCellFormula(abonCell + "*max(85-" + qualityCell + ",0)/100*" + penaltyCoefCell + "*$C$1*$C$2");

					cell.CellStyle = cs;
					currentRow = activeSheet.GetRow(dataStartRow + ++i);
				}
				int lastRow = activeSheet.LastRowNum;
				int lastColumn = activeSheet.GetRow(lastRow).LastCellNum;
				string summFirst = CellReference.ConvertNumToColString(penaltyColumnIndex) + (dataStartRow + 1);
				string summLast = CellReference.ConvertNumToColString(penaltyColumnIndex) + (lastRow);
				currentRow = activeSheet.GetRow(3);
				var doubleCellStyle = workbook.CreateCellStyle();
				doubleCellStyle.DataFormat = doubleFormat;
				var penaltyCell = currentRow.CreateCell(2);
				penaltyCell.CellStyle=doubleCellStyle;
				penaltyCell.SetCellFormula("SUM(" + summFirst + ":" + summLast + ")");
				
				
				MemoryStream ms = new MemoryStream();
				workbook.Write(ms);

				ExcelPackage ep = new ExcelPackage(ms);
				ep.Workbook.Worksheets[1].Drawings[0].SetPosition(lastRow, 0,lastColumn+1,0);
				ep.Workbook.Worksheets[1].Drawings[0].SetSize(113);

				using (FileStream fileStreamMod = new FileStream(Path.Combine(path, modeFilename), FileMode.Create))
				{
					ep.SaveAs(fileStreamMod);
				}
				workbook.Close();
				return modeFilename;
			}

		}

		private string GetPriceFromNameAndClass(string nameOps, string opsClass)
		{
			Dictionary<int, string> prices = new Dictionary<int, string>
			{
				{ 101,"203300"},{ 102,"66300"},{ 103,"49100"},{ 104,"35000"},
							{ 11,"58650"},{ 12,"30100"},{ 13,"17900"},{ 14,"11700"},{ 15,"34910"},
							{ 1,"66950"},{ 2,"44990"},{ 3,"13640"},{ 4,"5130"},{ 5,"3190"}
			};
			int myOpsClass = 0;
			if (int.TryParse(opsClass, out myOpsClass))
			{
				Regex regex = new Regex("отделение\\W+почтовой\\W+связи");
				if (regex.IsMatch(nameOps.ToLower()))
				{
					return prices[myOpsClass];
				}
				if (nameOps.ToLower().Contains("почтамт"))
				{
					return prices[myOpsClass+10];
				}
				regex = new Regex("участок\\W+курьерской\\W+доставки");
				if (regex.IsMatch(nameOps.ToLower()))
				{
					return prices[15];
				}
				if (nameOps.ToLower().Contains("филиал"))
					return prices[myOpsClass + 100];
				return "0";
			}
			return "0";
		}

		public Tuple<bool, string> GetPenalty(string passkey, string period)
		{
			DateTime currentPeriod = DateTime.Parse(period);
			int currentPeriodYear = currentPeriod.Year;
			int currentPeriodMonth = currentPeriod.Month;
			DateTime periodStart = new DateTime(currentPeriodYear, currentPeriodMonth, 1);
			DateTime periodFinish = new DateTime(currentPeriodYear, currentPeriodMonth,
				DateTime.DaysInMonth(currentPeriodYear, currentPeriodMonth)).AddDays(1).AddSeconds(-1);
			string filename = String.Format("report({0:yyyy-MM-dd HH_mm_ss}).xlsx", DateTime.Now);
			string reportRequestUrl =
			//$"https://support.russianpost.ru/sd/services/rest/create-m2m/report%24rep5167/?accessKey={passkey}";
			$"https://support.russianpost.ru/sd/services/rest/create-m2m/report%24rep5167/%7BcheckInTime:%22{periodStart:yyyy.MM.dd}%20{periodStart:HH:mm}%22,periodDate:%22{periodFinish:yyyy.MM.dd}%20{periodFinish:HH:mm}%22%7D?accessKey={passkey}";

			var data = DownloadString(reportRequestUrl);
			if (!data.Item1)
			{
				return new Tuple<bool, string>(false, "Ошибка. Не удалось запросить формирование отчета.");
				//DownloadLinkLabel.Text = "Ошибка. Не удалось запросить формирование отчета. Подробности в логах";
				//return;
			}
			JavaScriptSerializer javaScriptSerializer = new JavaScriptSerializer();
			var o = (dynamic)javaScriptSerializer.DeserializeObject(data.Item2);
			string report_id = null;
			try
			{
				report_id = o["UUID"];
			}
			catch (Exception ex)
			{
				MyLogger.GetInstance().Error("Ошибка парсинга ответа сервера Naumen на запрос отчета. В ответе не найдено поле UUID");
			}
			if (string.IsNullOrEmpty(report_id))
			{
				return new Tuple<bool, string>(false, "Ошибка. Не удалось запросить id файлa отчета.");
				MyLogger.GetInstance().Error("Ошибка парсинга ответа сервера Naumen на запрос отчета. Пустой UUID");
				//DownloadLinkLabel.Text = "Ошибка. Не удалось запросить id файлa отчета. Подробности в логах";
				//return;
			}
			report_id = report_id.Split('$')[1];
			//string report_id = "39942531";
			bool reportCheckResult = false;
			int reportCheckCount = 0;
			string reportCheckUrl =
				$"https://support.russianpost.ru/sd/services/rest/get/report%24{report_id}?accessKey={passkey}";
			while (!reportCheckResult && reportCheckCount < 60)
			{
				var checkResult = CheckReportIsReady(reportCheckUrl);
				if (checkResult.Item1)
					reportCheckResult = true;
				else
				{
					reportCheckCount++;
					System.Threading.Thread.Sleep(5000);
				}
			}
			if (reportCheckResult)
			{
				string getFileUrl = $"https://support.russianpost.ru/sd/services/rest/get-file/file%24{fileId}?accessKey={passkey}";
				byte[] excelData = DownloadData(getFileUrl);
				if (!Directory.Exists(HttpContext.Current.Server.MapPath("~/Storage/Reports/")))
					Directory.CreateDirectory(HttpContext.Current.Server.MapPath("~/Storage/Reports/"));
				File.WriteAllBytes(HttpContext.Current.Server.MapPath("~/Storage/Reports/" + filename), excelData);
				var serverRoot = HttpContext.Current.Request.Url.AbsoluteUri.Replace(HttpContext.Current.Request.Url.PathAndQuery, "");
				return new Tuple<bool, string>(true, "get-file/" + filename);
			}
			else
			{
				MyLogger.GetInstance().Warn("Таймаут формирования отчет Naumen");
				return new Tuple<bool, string>(false, "Таймаут формирования отчет Naumen");
				//DownloadLinkLabel.Text = "Не удалось выполнить формирование отчета Naumen. Попробуйте сформировать заново";

			}
		}

		public IndexModule()
		{
			Get["/"] = p =>
			{
				int startMonth = 10;
				int startYear = 2016;
				DateTime todayDate = DateTime.Today;
				int selectedIndex = 0;
				var monthList = new List<SelectListItem>();
				for (int i = 0; i < 20; i++)
				{
					int currentMonth = (startMonth + i) % 12 + 1;
					int currentYear = startYear + (startMonth + i) / 12;
					DateTime currentDate = new DateTime(currentYear, currentMonth, 1);
					monthList.Add(new SelectListItem()
					{
						Value = currentDate.ToString(),
						Text = currentDate.ToString("MMMM yyyy", System.Globalization.CultureInfo.GetCultureInfo("ru-RU"))
					});
					if (currentMonth == todayDate.Month && currentYear == todayDate.Year)
						selectedIndex = i;
				}
				if (todayDate.Day >= 1 && todayDate.Day <= DateTime.DaysInMonth(todayDate.Year, todayDate.Month) / 2)
					selectedIndex--;
				monthList[selectedIndex].Selected = true;
				IPAddress clientIpAddress = IPAddress.Parse(Request.UserHostAddress);
				var ipBytes = clientIpAddress.GetAddressBytes();
				RegionsModel regionsModel = new RegionsModel();
				regionsModel.Deserialize(HttpContext.Current.Server.MapPath("~/regions.xml"));
				RegionModel rm = null;
				foreach (RegionModel item in regionsModel.RegionModels)
				{
					if (item.IpSecondByte == ipBytes[1])
						rm = item;
				}
				if (rm!= null)
					return View["index", new { list = monthList, passkey = rm.Passkey, regionName = rm.Name }];
				else
					return View["index", new { list = monthList, passkey = "", regionName = "" }];
			};
			Post["/"] = p =>
			{
				string passkey = Request.Form["passkey"];
				string period = Request.Form["period"];
				var penaltyResult = GetPenalty(passkey, period);
				return Response.AsJson(new { result = penaltyResult.Item1, resultText = penaltyResult.Item2 });
			};
			Get["/get-file/{filename}"] = p =>
			{
				string realFilename = HttpContext.Current.Server.MapPath($"~/Storage/Reports/{p.filename}");
				if (File.Exists(realFilename))
					return Response.AsFile($"Storage/reports/{p.filename}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

				return new HttpNotFoundResult();
			};

			Get["/Views/index.cshtml"] = p =>
			{
				return Response.AsRedirect("/");
			};
			Get["/views/get.cshtml"] = p =>
			{
				return View["get", null];
			};
			Post["/views/get.cshtml"] = p =>
			{
				PenaltyModel penaltyModel = new PenaltyModel()
				{
					RegionName = passkeys[67].Item2,
					RegionCoeff = 0.97,
					EquipCoeff = 0.9904
				};
				string filename = "report(2017-01-30 10_58_15).xlsx";
				string modFilename = ModifyNaumenReport(HttpContext.Current.Server.MapPath("~/Storage/Reports/" + filename), penaltyModel);
				return Response.AsRedirect($"/get-file/{modFilename}");
			};
		}
	}
}