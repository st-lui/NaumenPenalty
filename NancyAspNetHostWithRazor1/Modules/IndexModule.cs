using System;
using Nancy;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Nancy.Json;
using Nancy.Responses;

namespace NancyAspNetHostWithRazor1.Modules
{
	public class IndexModule : NancyModule
	{
		string fileId;
		string resultText;
		Dictionary<byte, Tuple<string, string>> passkeys = new Dictionary<byte, Tuple<string, string>> {
			{ 67, new Tuple<string, string>("166470ff-beb7-48a3-b2a8-540d1b7c26dc", "Алтайский край") },
			{ 56, new Tuple<string, string>("b2e9acf7-09cb-4e72-8d66-588cbcfd8a30", "Республика Хакасия") } };

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
				File.WriteAllBytes(HttpContext.Current.Server.MapPath("~/Storage/Reports/" + filename), excelData);
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
				Tuple<string,string> model = null;
				if (passkeys.ContainsKey(ipBytes[1]))
					model = passkeys[ipBytes[1]];
				if (model!=null)
					return View["index", new { list = monthList, passkey=model.Item1,regionName=model.Item2 }];
				else
					return View["index", new { list = monthList, passkey = "", regionName = ""}];
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
				return Response.AsFile("Storage/site.css", "text/plain");
			};
		}
	}
}