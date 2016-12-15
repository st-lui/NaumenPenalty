using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Common;
using System.IO;
using NaumenPenalty.DataManage;
using NaumenPenalty.Model;
using Dapper;
using Dapper.Mapper;
using NaumenPenalty.Configuration;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace NaumenPenalty
{
	class Program
	{
		static IEnumerable<ServiceCall> r;
		static IEnumerable<Location> opses;
		static string idxPost;
		static int Main(string[] args)
		{

			if (args.Length < 2)
			{
				Console.WriteLine("отсутствуют обязательные параметры запуска");
				Console.WriteLine("- индекс филиала");
				Console.WriteLine("- дата периода");
				return 1;
			}
			idxPost = args[0];
			DateTime currentPeriod = DateTime.Parse(args[1]);
			int currentPeriodYear = currentPeriod.Year;
			int currentPeriodMonth = currentPeriod.Month;
			DateTime periodStart = new DateTime(currentPeriodYear, currentPeriodMonth, 1);
			DateTime periodFinish = new DateTime(currentPeriodYear, currentPeriodMonth,
				DateTime.DaysInMonth(currentPeriodYear, currentPeriodMonth)).AddDays(1).AddSeconds(-1);

			Mappings.RegisterDapperMappings();

			using (DbConnection connection = ConnectionFactory.CreateConnection())
			{

				//postoffice$typecode,postoffice$postalclass,parent

				string recursiveQuery =
					$@"with recursive tree 
(id,postoffice$postcode,postoffice$shortname,postoffice$typecode,postoffice$postalclass,parent,level) as (select id,postoffice$postcode,postoffice$shortname,postoffice$typecode,postoffice$postalclass,parent,0 
from tbl_location 
where postoffice$postcode =@idx 
union all 
select tbl_location.id, tbl_location.postoffice$postcode, tbl_location.postoffice$shortname,tbl_location.postoffice$typecode,tbl_location.postoffice$postalclass,tbl_location.parent,tree.level + 1 
from tbl_location,tree where tree.id = tbl_location.parent and removed=false) 
select id, postoffice$postcode,postoffice$shortname,postoffice$typecode,postoffice$postalclass,parent from tree Location";
				try
				{
					opses = connection.Query<Location>(recursiveQuery, new { idx = idxPost });
					List<long> postOfficeIds = new List<long>();
					string ids = "";
					foreach (var location in opses)
					{
						ids = ids + location.Id + ",";
					}
					ids = ids.Trim(',');
					string query =
$@"select ServiceCall.id
,ServiceCall.removed
,ServiceCall.agreement_id
,ServiceCall.registration_date
,ServiceCall.deadlinetime
,ServiceCall.mass_problem
,ServiceCall.number_
,ServiceCall.resolutiontime
,ServiceCall.state
,ServiceCall.statestarttime
,ServiceCall.clientou_id
,ServiceCall.timezone_id
,Agreement.id
,Agreement.title
,NaumenTimeZone.id
,NaumenTimeZone.code
,ClientOu.id
,ClientOu.location
,Location.id
,Location.postoffice$postcode
,Location.postoffice$shortname
from tbl_servicecall ServiceCall, tbl_agreement Agreement,tbl_ou ClientOu,tbl_location Location, tbl_timezone NaumenTimeZone
where ServiceCall.removed=false
and ServiceCall.agreement_id = Agreement.id 
and ServiceCall.timezone_id = NaumenTimeZone.id
and ServiceCall.registration_date>='{periodStart:yyyy-MM-dd}' 
and ServiceCall.registration_date<='{periodFinish:yyyy-MM-dd}'
and ServiceCall.clientou_id is not null
and ServiceCall.clientou_id = ClientOu.id
and ClientOu.location is not null
and ClientOu.location = Location.id
and Location.id in ({ids})";
					// порядок классов должен совпадать с порядком выбора id в запросе
					r = connection.Query<ServiceCall, Agreement, NaumenTimeZone, Ou, Location>(query);
					var excelData = CreatePenaltiesReport(currentPeriod.Year, currentPeriod.Month, "template", int.Parse(idxPost), false);
					File.WriteAllBytes("pen.xls", excelData);
					Console.WriteLine("DONE");
					Console.ReadLine();
				}
				catch (Exception e)
				{
					Console.WriteLine(e.ToString());
					if (e.InnerException != null)
					{
						Console.WriteLine("Inner exception");
						Console.WriteLine(e.InnerException.ToString());
					}
					Console.ReadLine();
				}
			}
			return 0;
		}

		/// <summary>
		/// Структура для хранения данных о заявке для формирования второго листа отчета по штрафам
		/// </summary>
		public struct RequestData
		{
			public DateTime? closeDate;
			public DateTime? actCloseDate;
			public string penaltyType;
			public string comment;
			public long norma;
			public DateTime normativeDate;

			public RequestData(DateTime? closeDate, DateTime? actCloseDate, string penaltyType, string comment, long norma, DateTime normativeDate)
			{

				this.closeDate = closeDate;
				this.actCloseDate = actCloseDate;
				this.penaltyType = penaltyType;
				this.comment = comment;
				this.norma = norma;
				this.normativeDate = normativeDate;
			}
		}

		public static byte[] CreatePenaltiesReport(int currentPeriodYear, int currentPeriodMonth, string templateDirectory, int postId, bool considerActs)
		{
			var configManager = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
			using (FileStream fs = new FileStream(Path.Combine(templateDirectory, "penalties.xls"), FileMode.Open))
			{
				HSSFWorkbook workbook = new HSSFWorkbook(fs);
				HSSFSheet activeSheet = (HSSFSheet)workbook.GetSheetAt(0);
				int startRow = activeSheet.FirstRowNum + 10;
				int i = 0, rowIndex = startRow;
				DateTime periodStart = new DateTime(currentPeriodYear, currentPeriodMonth, 1);
				DateTime periodFinish = new DateTime(currentPeriodYear, currentPeriodMonth,
					DateTime.DaysInMonth(currentPeriodYear, currentPeriodMonth)).AddDays(1).AddSeconds(-1);
				var row1 = activeSheet.GetRow(2);
				row1.Cells[1].SetCellValue(string.Format("{0:d} - {1:d}", periodStart, periodFinish));
				var feeConfigSection = (FeeConfigSection) configManager.GetSection("Fees");
				FeeCollection feeCollection = null;
				if (feeConfigSection != null)
					feeCollection = (FeeCollection) feeConfigSection.Items;
				var penaltyMultipliers = (PenaltyMultiplierConfigSection)configManager.GetSection("PenaltyMultipliers");
				if (penaltyMultipliers != null)
				{
					var penaltyMultiplier = (PenaltyMultiplierElement)penaltyMultipliers.Items[postId];
					if (penaltyMultiplier != null)
					{
						activeSheet.GetRow(0).Cells[1].SetCellValue(penaltyMultiplier.RegionalMultiplier);
						activeSheet.GetRow(1).Cells[1].SetCellValue(penaltyMultiplier.EquipmentMultiplier);
					}
				}
				else
				{
					activeSheet.GetRow(0).Cells[1].SetCellValue(0.5);
					activeSheet.GetRow(1).Cells[1].SetCellValue(0.63);
				}

				//int softwareServiceId = ServiceController.SelectIdByText("Обновление, установка и настройка ПО");
				int softwareServiceId = 0;
				List<KeyValuePair<ServiceCall, RequestData>> requestDatas = new List<KeyValuePair<ServiceCall, RequestData>>();
				foreach (Location ops in opses)
				{
					//entities.Attach(south);
					var reqs = r.Where(x => x.ClientOu.Location.Id == ops.Id).ToList();
					int заявштр1 = 0;
					int заявштр2 = 0;
					int заявштр3 = 0;
					int заявштр4 = 0;
					int заявсвоевр1 = 0;
					int заявсвоевр2 = 0;

					double качествоSLA = 85;

					List<ServiceCall> заявштр1List = new List<ServiceCall>();
					List<ServiceCall> заявштр2List = new List<ServiceCall>();
					List<ServiceCall> заявштр3List = new List<ServiceCall>();
					List<ServiceCall> заявштр4List = new List<ServiceCall>();
					List<ServiceCall> заявсвоевр1List = new List<ServiceCall>();
					List<ServiceCall> заявсвоевр2List = new List<ServiceCall>();

					getOpsPenalties(reqs, periodStart, periodFinish, softwareServiceId, ref заявштр1, ref заявштр2, ref заявштр3, ref заявштр4, ref заявсвоевр1, ref заявсвоевр2,
						заявштр1List, заявштр2List, заявштр3List, заявштр4List, заявсвоевр1List, заявсвоевр2List, requestDatas);

					if (заявсвоевр1 + заявсвоевр2 + заявштр1 + заявштр2 + заявштр3 + заявштр4 > 0)
					{
						double качествоФакт = 100.0 -
											  1.0 * (заявштр1 + заявштр2 + заявштр3 + заявштр4) /
											  (заявсвоевр1 + заявсвоевр2 + заявштр1 + заявштр2 + заявштр3 + заявштр4) * 100;

						//entities.Detach(south);
						var row = activeSheet.CreateRow(rowIndex);
						int startColumnIndex = row.FirstCellNum + 1;
						for (int col = 0; col < 16; col++)
							row.CreateCell(startColumnIndex + col);
						row.Cells[startColumnIndex].SetCellValue(ops.PostCode + " " + ops.ShortName);
						row.Cells[startColumnIndex + 1].SetCellValue(заявштр1);
						row.Cells[startColumnIndex + 2].SetCellValue(заявштр2);
						row.Cells[startColumnIndex + 3].SetCellValue(заявштр3);
						row.Cells[startColumnIndex + 4].SetCellValue(заявштр4);
						row.Cells[startColumnIndex + 5].SetCellValue(заявсвоевр1);
						row.Cells[startColumnIndex + 6].SetCellValue(заявсвоевр2);
						row.Cells[startColumnIndex + 7].SetCellValue(качествоФакт);
						//writer.Write(
						//    "ОПС:{0}\tЗаявШТР1:{1}\tЗаявШТР2:{2}\tЗаявШТР3:{3}\tЗаявШТР4:{4}\tЗаявСВОЕВР1:{5}\tЗаявСВОЕВР2:{6}\tКачествоФакт:{7}\t",
						//    ops.name_post + " " + ops.idx + " " + ops.name_ops, заявштр1, заявштр2, заявштр3, заявштр4, заявсвоевр1,
						//    заявсвоевр2, качествоФакт);

						double качестоРасч = Math.Max(качествоSLA - качествоФакт, 0);
						double кштр = 0;
						double качествоФактRounded = Math.Round(качествоФакт);
						if (качествоФактRounded < 73)
							кштр = 2;
						else if (качествоФактRounded >= 73 && качествоФактRounded <= 76)
							кштр = 1.5;
						else if (качествоФактRounded >= 76 && качествоФактRounded <= 85)
							кштр = 1;
						double СТаб = 0;
						string classString = "";
						//филиал
						int opsTypeId = 0;
						if (!ops.Parent.HasValue)
						{
							switch (ops.PostalClass)
							{
								case 1:
									opsTypeId = 101;
									break;
								case 2:
									opsTypeId = 102;
									break;
								case 3:
									opsTypeId = 103;
									break;
								case 4:
									opsTypeId = 104;
									break;
								default:
									classString = "Не указана категория филиала";
									break;
							}
						}
						//почтамт
						else if (ops.Type == "РП" || ops.Type == "МРП" || ops.Type == "ГП" || ops.Type == "Почтамт")
						{
							switch (ops.PostalClass)
							{
								case 1:
									opsTypeId = 11;
									break;
								case 2:
									opsTypeId = 12;
									break;
								case 3:
									opsTypeId = 13;
									break;
								case 4:
									opsTypeId = 14;
									break;
								default:
									classString = "Не указана категория почтамта";
									break;
							}
						}
						//опс
						else if (ops.Type == "СОПС" || ops.Type == "ГОПС" || ops.Type == "ПОПС")
						{
							switch (ops.PostalClass)
							{
								case 1:
									opsTypeId = 1;
									break;
								case 2:
									opsTypeId = 2;
									break;
								case 3:
									opsTypeId = 3;
									break;
								case 4:
									opsTypeId = 4;
									break;
								case 5:
									opsTypeId = 5;
									break;
								default:
									classString = "Не  указан класс ОПС";
									break;
							}
						}
						//укд
						else if (ops.Type == "УКД")
							opsTypeId = 15;
						else
							classString = "Не указан тип объекта";
						row.Cells[startColumnIndex + 8].SetCellValue(кштр);
						if (opsTypeId > 0)
						{
							FeeElement fee = (FeeElement) feeCollection?[opsTypeId];
							if (fee != null)
							{
								classString = fee.Description;
								СТаб = fee.FeeValue;
							}
						}
						row.Cells[startColumnIndex + 9].SetCellValue(classString);
						row.Cells[startColumnIndex + 10].SetCellValue(СТаб);
						string abonCell = CellReference.ConvertNumToColString(startColumnIndex + 11 - 1) + (row.RowNum + 1);
						string penaltyCoefCell = CellReference.ConvertNumToColString(startColumnIndex + 11 - 3) + (row.RowNum + 1);
						string qualityCell = CellReference.ConvertNumToColString(startColumnIndex + 11 - 4) + (row.RowNum + 1);
						row.Cells[startColumnIndex + 11].SetCellFormula(abonCell + "*max(85-" + qualityCell + ",0)/100*" + penaltyCoefCell + "*$B$1*$B$2");
						//row.Cells[startC	olumnIndex+11].r;
						rowIndex++;
					}
				}
				requestDatas.Sort(
					(x, y) => x.Key.RegistrationDate.CompareTo(y.Key.RegistrationDate));
				ISheet secondSheet = workbook.GetSheetAt(1);
				int firstrow = secondSheet.FirstRowNum + 1;
				for (int index = 0; index < requestDatas.Count; index++)
				{
					var requestData = requestDatas[index];
					IRow currentRow = secondSheet.CreateRow(firstrow + index);
					int firstcolumn = 0;
					//номер заявки
					currentRow.CreateCell(firstcolumn).SetCellValue(requestData.Key.Number);
					//ОПС
					currentRow.CreateCell(firstcolumn + 1).SetCellValue(requestData.Key.ClientOu.Location.PostCode + "_" + requestData.Key.ClientOu.Location.ShortName);
					//вид услуги
					currentRow.CreateCell(firstcolumn + 2).SetCellValue(requestData.Key.Agreement.Title);
					// статус
					currentRow.CreateCell(firstcolumn + 3).SetCellValue(requestData.Key.State);
					// дата заявки
					currentRow.CreateCell(firstcolumn + 4).SetCellValue(requestData.Key.RegistrationDate.Add(requestData.Key.NaumenTimeZone.TimeZoneInfo.BaseUtcOffset).ToString());
					//if (requestData.Value.actCloseDate.HasValue)
					//	ts = requestData.Value.actCloseDate.Value -
					//		 requestData.Key.DATECREATED.AddHours(requestData.Key.USER.POST.regions.timezone);
					//else
					//	if (requestData.Value.closeDate.HasValue)
					//	ts = requestData.Value.closeDate.Value -
					//	 requestData.Key.DATECREATED.AddHours(requestData.Key.USER.POST.regions.timezone);
					//время выполнения заявки
					TimeSpan workTs = new TimeSpan();
					if (requestData.Value.closeDate.HasValue)
						workTs = requestData.Value.closeDate.Value -
								 requestData.Key.RegistrationDate.Add(requestData.Key.NaumenTimeZone.TimeZoneInfo.BaseUtcOffset);
					else
						workTs = DateTime.Now - requestData.Key.RegistrationDate.Add(requestData.Key.NaumenTimeZone.TimeZoneInfo.BaseUtcOffset);
					currentRow.CreateCell(firstcolumn + 5).SetCellValue(workTs.Days + " д. " + workTs.Subtract(new TimeSpan(workTs.Days, 0, 0, 0)).ToString(@"hh\:mm\:ss"));
					// дата закрытия
					currentRow.CreateCell(firstcolumn + 6).SetCellValue(requestData.Value.closeDate.HasValue ? requestData.Value.closeDate.Value.ToString() : string.Empty);
					//// дата закрытия + наличие акта
					//currentRow.CreateCell(firstcolumn + 8).SetCellValue(requestData.Value.actCloseDate.HasValue ? requestData.Value.actCloseDate.Value.ToString() : string.Empty);
					// срок ремонта
					// дни
					TimeSpan ts = TimeSpan.FromMilliseconds(requestData.Key.ResolutionTime);
					currentRow.CreateCell(firstcolumn + 7).SetCellValue(ts.Days + " д. " + ts.Subtract(new TimeSpan(ts.Days, 0, 0, 0)).ToString(@"hh\:mm\:ss"));

					// комментарий
					currentRow.CreateCell(firstcolumn + 8).SetCellValue(requestData.Value.comment);
					// планируемая дата закрытия
					currentRow.CreateCell(firstcolumn + 9).SetCellValue(requestData.Value.normativeDate.ToString());
					//штрафы
					currentRow.CreateCell(firstcolumn + 10).SetCellValue(requestData.Value.penaltyType);
				}
				row1 = activeSheet.GetRow(activeSheet.FirstRowNum + 4);
				string summFirst = CellReference.ConvertNumToColString(11) + (startRow + 1);
				string summLast = CellReference.ConvertNumToColString(11) + (rowIndex);
				row1.Cells[1].SetCellFormula("SUM(" + summFirst + ":" + summLast + ")");
				MemoryStream ms = new MemoryStream();
				workbook.Write(ms);
				ms.Flush();
				var buffer = ms.GetBuffer();
				ms.Dispose();
				ms.Close();
				return buffer;
			}
		}

		public static void getOpsPenalties(List<ServiceCall> reqs, DateTime periodStart, DateTime periodFinish, int softwareServiceId, ref int заявштр1, ref int заявштр2, ref int заявштр3, ref int заявштр4, ref int заявсвоевр1, ref int заявсвоевр2, List<ServiceCall> заявштр1List, List<ServiceCall> заявштр2List, List<ServiceCall> заявштр3List, List<ServiceCall> заявштр4List, List<ServiceCall> заявсвоевр1List, List<ServiceCall> заявсвоевр2List, List<KeyValuePair<ServiceCall, RequestData>> requestDatas)
		{
			foreach (var request in reqs)
			{
				TimeZoneInfo timeZoneInfo = request.NaumenTimeZone.TimeZoneInfo;
				DateTime requestDate = request.RegistrationDate.Add(timeZoneInfo.BaseUtcOffset);
				string message;
				var norma = request.ResolutionTime;
				if (norma == null)
					continue;
				DateTime normativeDate = request.DeadLineTime.Add(timeZoneInfo.BaseUtcOffset);
				// если заявка  не закрыта
				if (request.State.ToLower() != "closed" && request.State.ToLower() != "resolved")
				{
					// неисполненные с целевой датой в прошлом периоде
					if (normativeDate < periodStart)
					{
						// не включать если настройка и обновление ПО
						if (request.ServiceId != softwareServiceId)
						{
							заявштр1++;
							заявштр1List.Add(request);
							requestDatas.Add(new KeyValuePair<ServiceCall, RequestData>(request,
								new RequestData(null, null, "ЗАЯВштр1", "Не выполнена.", norma, normativeDate)));
						}
					}
					else
					// неисполненные с целевой датой в текущем периоде
					if (periodStart <= normativeDate && normativeDate <= periodFinish)
					{
						заявштр2++;
						заявштр2List.Add(request);
						requestDatas.Add(new KeyValuePair<ServiceCall, RequestData>(request,
							new RequestData(null, null, "ЗАЯВштр2", "Не выполнена.", norma, normativeDate)));
					}
				}
				// если заявка закрыта
				else
				{
					DateTime realCloseDate = request.StateStartTime.Add(request.NaumenTimeZone.TimeZoneInfo.BaseUtcOffset);
					//DateTime closeDate = realCloseDate;
					//DateTime? actCloseDate = null;
					//if (request.act_date.HasValue)
					//{
					//	realCloseDate = request.act_date.Value.AddHours(request.USER.POST.regions.timezone);
					//	actCloseDate = realCloseDate;
					//}
					// закрытые в текущем периоде с целевой датой в прошлом периоде
					if (periodStart <= realCloseDate && realCloseDate <= periodFinish && normativeDate < periodStart)
					{
						// не включать если настройка и обновление ПО
						if (request.ServiceId != softwareServiceId)
						{
							заявштр3++;
							заявштр3List.Add(request);
							requestDatas.Add(new KeyValuePair<ServiceCall, RequestData>(request,
								new RequestData(realCloseDate, realCloseDate, "ЗАЯВштр3", "Выполнена не в срок.", norma, normativeDate)));
						}
					}
					else
					// закрытые в текущем периоде с целевой датой в текущем периоде с нарушением срока выполнения
					if (periodStart <= realCloseDate && realCloseDate <= periodFinish && periodStart <= normativeDate &&
							 normativeDate <= periodFinish &&
							 realCloseDate > normativeDate)
					{
						заявштр4++;
						заявштр4List.Add(request);
						requestDatas.Add(new KeyValuePair<ServiceCall, RequestData>(request,
							new RequestData(realCloseDate, realCloseDate, "ЗАЯВштр4", "Выполнена не в срок.", norma, normativeDate)));
					}
					else
					// закрытые в текущем периоде с целевой датой в текущем периоде с выполнением в срок
					if (periodStart <= realCloseDate && realCloseDate <= periodFinish && periodStart <= normativeDate &&
							 normativeDate <= periodFinish &&
							 realCloseDate <= normativeDate)
					{
						заявсвоевр1++;
						заявсвоевр1List.Add(request);
						requestDatas.Add(new KeyValuePair<ServiceCall, RequestData>(request,
							new RequestData(realCloseDate, realCloseDate, "ЗАЯВсвоевр1", "Выполнена в срок.", norma, normativeDate)));
					}
					else
					// закрытые в текущем периоде с целевой датой в следующем периоде
					if (periodStart <= realCloseDate && realCloseDate <= periodFinish && normativeDate > periodFinish &&
							 realCloseDate <= normativeDate)
					{
						заявсвоевр2++;
						заявсвоевр2List.Add(request);
						requestDatas.Add(new KeyValuePair<ServiceCall, RequestData>(request,
							new RequestData(realCloseDate, realCloseDate, "ЗАЯВсвоевр2", "Выполнена в срок.", norma, normativeDate)));
					}
				}
			}
		}

	}
}
