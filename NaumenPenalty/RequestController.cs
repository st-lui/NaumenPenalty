using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Objects;
using System.Data.Objects.SqlClient;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web;
using HD_CLASSES;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using System.Linq.Dynamic;
using Microsoft.Exchange.WebServices.Data;
using NPOI.SS.Util;
using ThreadState = System.Threading.ThreadState;
//using Word = Microsoft.Office.Interop.Word;

namespace DataAccess
{
	[System.ComponentModel.DataObject]
	public class RequestController
	{
		private List<int> postIdTree;

		public RequestController()
		{

		}

		public static void Save(REQUEST r)
		{
			using (hdEntities entities = new hdEntities())
			{
				entities.REQUEST.AddObject(r);
				entities.SaveChanges();
				entities.LoadProperty(r, x => x.POST);
			}
		}

		public static void Save(REQUEST r, List<DEVICEINFO> deviceinfos)
		{
			using (hdEntities entities = new hdEntities())
			{
				entities.REQUEST.AddObject(r);
				entities.SaveChanges();
				for (int i = 0; i < deviceinfos.Count; i++)
				{
					entities.Attach(deviceinfos[i]);
					entities.Detach(deviceinfos[i].OPS);
				}
				for (int i = 0; i < deviceinfos.Count; i++)
				{
					r.DEVICEINFOS.Add(deviceinfos[i]);
				}
				entities.SaveChanges();
				entities.LoadProperty(r, x => x.POST);
			}
		}

		public static void Update(REQUEST r, List<DEVICEINFO> deviceinfos)
		{
			using (hdEntities entities = new hdEntities())
			{
				REQUEST req = entities.REQUEST.Single(x => x.ID == r.ID);
				for (int i = 0; i < deviceinfos.Count; i++)
				{
					int dev_id = deviceinfos[i].ID;
					DEVICEINFO dev = entities.DEVICEINFO.Single(x => x.ID == dev_id);
					req.DEVICEINFOS.Add(dev);
				}
				entities.SaveChanges();
				//entities.LoadProperty(r, x => x.POST);
			}
		}

		public static void RemoveDeviceInfo(REQUEST r, DEVICEINFO deviceinfo)
		{
			using (hdEntities entities = new hdEntities())
			{
				entities.Attach(r);
				r.DEVICEINFOS.Remove(deviceinfo);
				entities.SaveChanges();
				entities.Detach(deviceinfo);
			}
		}

		public void SaveRange(List<REQUEST> range)
		{
			using (hdEntities entities = new hdEntities())
			{
				foreach (var request in range)
				{
					entities.REQUEST.AddObject(request);
				}
				entities.SaveChanges();
			}
		}

		public static void Update(hdEntities entities, REQUEST request)
		{
			entities.SaveChanges();
		}

		public static void Update(REQUEST request)
		{
			using (hdEntities entities = new hdEntities())
			{
				entities.Attach(request);
				//entities.Entry(request).State = EntityState.Modified;
				entities.ObjectStateManager.ChangeObjectState(request, EntityState.Modified);
				entities.SaveChanges();
			}
		}

		public static List<REQUEST> Find()
		{
			using (hdEntities entities = new hdEntities())
			{
				return entities.REQUEST.ToList();
			}
		}

		/// <summary>
		/// Выбор всех заявок для POST
		/// </summary>
		public IList<REQUEST> FindByPost(int postId)
		{
			using (hdEntities entities = new hdEntities())
			{
				//var requests = from r in entities.REQUEST where r.POST_ID == postId select new RequestSelectResult(){ POSTNAME = r.POST.NAME, USERNAME = r.USER.NAME, DATECREATED = r.DATECREATED, PRIORITYNUMBER = r.PRIORITY.NUMBER, OPS_ID = r.OPS_ID, TEXT = r.TEXT,ID=r.ID,SOLVED=(r.DATESOLVED!=null && r.NUMBER!=null),DATESOLVED = r.DATESOLVED,NUMBER = r.NUMBER};
				postIdTree = new List<int>();
				searchPostIdTree(postId);
				var requests =
					entities.REQUEST.Include("POST")
						.Include("USER")
						.Include("BF_STATUS")
						.Include("DEVICEINFOS")
						.Include("USER.POST.regions")
						.Include("DEVICEINFOS.OPS")
						.Include("services")
						.Where(x => postIdTree.Contains(x.POST_ID))
						.ToList();
				return requests;
			}
		}

		public static REQUEST FindByMaykor(string maykorNumber)
		{
			using (hdEntities entities = new hdEntities())
			{
				REQUEST r =
					entities.REQUEST.Include("POST")
						.Include("USER")
						.Include("BF_STATUS")
						.Include("DEVICEINFOS")
						.Include("DEVICEINFOS.OPS")
						.Include("children")
						.Include("children.deviceinfos")
						.SingleOrDefault(x => x.BF_NUMBER == maykorNumber);
				//if (r!=null)
				//{
				//	foreach (var deviceinfo in r.DEVICEINFOS)
				//		deviceinfo.OPSReference.Load();
				//}
				return r;
			}
		}


		//<asp:Parameter Name="userId" Type="Int32"/>
		//            <asp:Parameter Name="dateCreatedStart" Type="DateTime"/>
		//            <asp:Parameter Name="dateCreatedEnd" Type="DateTime"/>
		//            <asp:Parameter Name="priorityId" Type="Int32"/>
		//            <asp:Parameter Name="opsId" Type="Int32"/>
		//            <asp:Parameter Name="solved" Type="Boolean"/>
		//            <asp:Parameter Name="dateSolvedStart" Type="DateTime"/>
		//            <asp:Parameter Name="dateSolvedEnd" Type="DateTime"/>
		//            <asp:Parameter Name="number" Type="Int32"/>

		//public IList<REQUEST> FindByPost(int postId, string sortExpression, Int32 userId, int postIdFilter,
		//    string opsId, Int32 solved, string deviceModel,
		//    int bfStatusId, string bfNumber, string deviceNumber)
		//{
		//    return FindByPost(postId, sortExpression, userId, postIdFilter, SqlDateTime.MinValue.Value,
		//        SqlDateTime.MaxValue.Value,
		//        opsId, solved, deviceModel, bfStatusId, bfNumber, deviceNumber);
		//}

		///// <summary>
		///// Выбор всех заявок с учетом фильтра
		///// </summary>
		//public IList<REQUEST> FindByPost(int postId, string sortExpression, int userId, int postIdFilter,
		//	DateTime dateCreatedStart, DateTime dateCreatedEnd, string opsId,
		//	int solved, string deviceModel, int bfStatusId, string bfNumber, string deviceNumber, string parentBfNumber, string text)
		//{
		//	postIdTree = new List<int>();
		//	// postId прилетает от пользователя
		//	searchPostIdTree(postId);
		//	if (String.IsNullOrWhiteSpace(sortExpression))
		//	{
		//		sortExpression = "DATECREATED desc";
		//	}
		//	var userList = new List<int>();
		//	if (userId > 0)
		//		userList.Add(userId);
		//	else
		//		userList.AddRange(UserController.SelectUsers(postId).Select(user => user.ID));
		//	if (postIdFilter > 0)
		//	{
		//		postIdTree.Clear();
		//		postIdTree.Add(postIdFilter);
		//	}
		//	var bfStatusList = new List<int>();
		//	if (bfStatusId > 0)
		//		bfStatusList.Add(bfStatusId);
		//	else
		//	{
		//		bfStatusList.AddRange(BfStatusController.List().Select(x => x.ID));
		//		bfStatusList.Add(0); // значение по умолчанию
		//	}
		//	opsId = opsId ?? string.Empty;
		//	deviceModel = deviceModel ?? string.Empty;
		//	deviceModel = deviceModel.ToLower();
		//	bfNumber = bfNumber ?? string.Empty;
		//	bfNumber = bfNumber.ToLower().Trim();
		//	deviceNumber = deviceNumber ?? string.Empty;
		//	deviceNumber = deviceNumber.ToLower().Trim();
		//	parentBfNumber = parentBfNumber ?? string.Empty;
		//	parentBfNumber = parentBfNumber.ToLower().Trim();
		//	text = text ?? string.Empty;
		//	text = text.ToLower().Trim();

		//	using (hdEntities entities = new hdEntities())
		//	{
		//		List<REQUEST> r = new List<REQUEST>();
		//		r = (from x in entities.REQUEST
		//			 let post_id = x.POST_ID
		//			 let author_id = x.AUTHOR_ID
		//			 let bf_status_id = x.BF_STATUS_ID.HasValue ? x.BF_STATUS_ID.Value : 0
		//			 let solvedvalue = x.SOLVED.HasValue && x.SOLVED.Value
		//			 let dev0 = x.DEVICEINFOS.FirstOrDefault()
		//			 let datecreated = x.DATECREATED
		//			 let timezone = x.USER.POST.regions.timezone
		//			 let bf_number = x.BF_NUMBER
		//			 let parent_bf_number = x.parent_id.HasValue ? x.parent.BF_NUMBER : "-"
		//			 let _text = x.TEXT.ToLower().Trim()
		//			 where
		//				postIdTree.Contains(post_id) && userList.Contains(author_id) &&
		//				//x.DEVICEINFOS.Any(y => postList.Contains(y.OPS.idx_post.Value)) &&
		//				//dev0.OPS.idx_post == postIdFilter &&
		//				//flag &&
		//				bfStatusList.Contains(bf_status_id) &&
		//				(solved == 0 || solved == 1 && solvedvalue ||
		//				 solved == 2 && !solvedvalue) &&
		//				EntityFunctions.AddHours(datecreated, timezone) >= dateCreatedStart &&
		//				EntityFunctions.AddHours(datecreated, timezone) <= dateCreatedEnd &&
		//				SqlFunctions.StringConvert((double?)(dev0.OPS.idx)).Contains(opsId) &&
		//				(dev0.type ?? string.Empty).ToLower().Contains(deviceModel) &&
		//				(dev0.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) &&
		//				(bf_number ?? string.Empty).ToLower().Contains(bfNumber) &&
		//				parent_bf_number.ToLower().Contains(parentBfNumber) &&
		//				_text.Contains(text)

		//			 //&&
		//			 //(x.DEVICEINFO == null && (x.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) ||
		//			 //x.DEVICEINFO != null &&
		//			 //(x.DEVICEINFO.DEVICENUMBER ?? String.Empty).ToLower().Contains(deviceNumber))
		//			 select x).ToList();
		//		return r;
		//	}
		//}
		/// <summary>
		/// Выбор диапазона заявок, фильтр включает даты
		/// </summary>
		public IList<REQUEST> FindByPost(int postId, string sortExpression, int userId, int postIdFilter,
										 DateTime dateCreatedStart, DateTime dateCreatedEnd, string opsId,
										 int solved, string deviceModel, int maximumRows, int startRowIndex,
										 int bfStatusId, string bfNumber, string deviceNumber, string parentBfNumber, string text)
		{
			postIdTree = new List<int>();
			searchPostIdTree(postId);
			if (String.IsNullOrWhiteSpace(sortExpression))
			{
				sortExpression = "DATECREATED desc";
			}
			//var userList = new List<int>();
			//if (userId > 0)
			//	userList.Add(userId);
			//else
			//	userList.AddRange(UserController.SelectUsers(postId).Select(user => user.ID));
			if (postIdFilter > 0)
			{
				postIdTree.Clear();
				postIdTree.Add(postIdFilter);
			}
			var bfStatusList = new List<int>();
			if (bfStatusId > 0)
				bfStatusList.Add(bfStatusId);
			else
			{
				bfStatusList.AddRange(BfStatusController.List().Select(x => x.ID));
				bfStatusList.Add(0); // значение по умолчанию
			}
			opsId = opsId ?? string.Empty;
			deviceModel = deviceModel ?? string.Empty;
			deviceModel = deviceModel.ToLower();
			bfNumber = bfNumber ?? string.Empty;
			bfNumber = bfNumber.ToLower().Trim();
			deviceNumber = deviceNumber ?? string.Empty;
			deviceNumber = deviceNumber.ToLower().Trim();
			parentBfNumber = parentBfNumber ?? string.Empty;
			parentBfNumber = parentBfNumber.ToLower().Trim();
			text = text ?? string.Empty;
			text = text.ToLower().Trim();
			using (hdEntities entities = new hdEntities())
			{
				List<REQUEST> r = new List<REQUEST>();
				if (maximumRows == 0 && startRowIndex == 0)
					r = (from x in entities.REQUEST
						 let post_id = x.POST_ID
						 let author_id = x.AUTHOR_ID
						 let bf_status_id = x.BF_STATUS_ID.HasValue ? x.BF_STATUS_ID.Value : 0
						 let solvedvalue = x.SOLVED.HasValue && x.SOLVED.Value
						 let dev0 = x.DEVICEINFOS.FirstOrDefault()
						 let datecreated = x.DATECREATED
						 let timezone = x.USER.POST.regions.timezone
						 let bf_number = x.BF_NUMBER
						 let parent_bf_number = x.parent_id.HasValue ? x.parent.BF_NUMBER : "-"
						 let _text = x.TEXT.ToLower().Trim()
						 where
							postIdTree.Contains(post_id) && (userId == 0 || userId != 0 && author_id == userId) &&
							bfStatusList.Contains(bf_status_id) &&
							(solved == 0 || solved == 1 && solvedvalue ||
							 solved == 2 && !solvedvalue) &&
							EntityFunctions.AddHours(datecreated, timezone) >= dateCreatedStart &&
							EntityFunctions.AddHours(datecreated, timezone) <= dateCreatedEnd &&
							SqlFunctions.StringConvert((double?)(dev0.OPS.idx)).Contains(opsId) &&
							(dev0.type ?? string.Empty).ToLower().Contains(deviceModel) &&
							(dev0.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) &&
							(bf_number ?? string.Empty).ToLower().Contains(bfNumber) &&
							parent_bf_number.ToLower().Contains(parentBfNumber) &&
							_text.Contains(text)
						 //&&
						 //(x.DEVICEINFO == null && (x.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) ||
						 //x.DEVICEINFO != null &&
						 //(x.DEVICEINFO.DEVICENUMBER ?? String.Empty).ToLower().Contains(deviceNumber))
						 select x).OrderBy(sortExpression).ToList();
				else
					r = (from x in entities.REQUEST//.Include("POST").Include("OPS").Include("USER").Include("USER.POST").Include("USER.POST.regions").Include("DEVICEINFOS")
						 let post_id = x.POST_ID
						 let author_id = x.AUTHOR_ID
						 let bf_status_id = x.BF_STATUS_ID.HasValue ? x.BF_STATUS_ID.Value : 0
						 let solvedvalue = x.SOLVED.HasValue && x.SOLVED.Value
						 let dev0 = x.DEVICEINFOS.FirstOrDefault()
						 let datecreated = x.DATECREATED
						 let timezone = x.USER.POST.regions.timezone
						 let bf_number = x.BF_NUMBER
						 let parent_bf_number = x.parent_id.HasValue ? x.parent.BF_NUMBER : "-"
						 let _text = x.TEXT.ToLower().Trim()
						 where
							postIdTree.Contains(post_id) && (userId == 0 || userId != 0 && author_id == userId) &&
							bfStatusList.Contains(bf_status_id) &&
							(solved == 0 || solved == 1 && solvedvalue ||
							 solved == 2 && !solvedvalue) &&
							EntityFunctions.AddHours(datecreated, timezone) >= dateCreatedStart &&
							EntityFunctions.AddHours(datecreated, timezone) <= dateCreatedEnd &&
							SqlFunctions.StringConvert((double?)(dev0.OPS.idx)).Contains(opsId) &&
							(dev0.type ?? string.Empty).ToLower().Contains(deviceModel) &&
							(dev0.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) &&
							(bf_number ?? string.Empty).ToLower().Contains(bfNumber) &&
							parent_bf_number.ToLower().Contains(parentBfNumber) &&
							_text.Contains(text)
						 //&&
						 //(x.DEVICEINFO == null && (x.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) ||
						 //x.DEVICEINFO != null &&
						 //(x.DEVICEINFO.DEVICENUMBER ?? String.Empty).ToLower().Contains(deviceNumber))
						 select x).OrderBy(sortExpression).Skip(startRowIndex).Take(maximumRows).ToList();

				foreach (var request in r)
				{
					if (!request.POSTReference.IsLoaded)
						entities.LoadProperty(request, "POST");
					if (!request.BF_STATUSReference.IsLoaded)
						entities.LoadProperty(request, "BF_STATUS");
					if (!request.DEVICEINFOS.IsLoaded)
						entities.LoadProperty(request, "DEVICEINFOS");
					if (!request.DEVICEINFOS.First().OPSReference.IsLoaded)
						entities.LoadProperty(request.DEVICEINFOS.First(), x => x.OPS);
					if (!request.USERReference.IsLoaded)
						entities.LoadProperty(request, "USER");
					if (!request.USER.POSTReference.IsLoaded)
						entities.LoadProperty(request.USER, "POST");
					if (!request.USER.POST.regionsReference.IsLoaded)
						entities.LoadProperty(request.USER.POST, "regions");
					if (!request.servicesReference.IsLoaded)
						entities.LoadProperty(request, "services");
				}
				return r;
			}

		}
		/// <summary>
		/// Выбор диапазона заявок, фильтр не включает даты
		/// </summary>
		public IList<REQUEST> FindByPost(int postId, string sortExpression, int userId, int postIdFilter,
										  string opsId, int solved, string deviceModel, int maximumRows, int startRowIndex,
										 int bfStatusId, string bfNumber, string deviceNumber, string parentBfNumber, string text)
		{
			postIdTree = new List<int>();
			searchPostIdTree(postId);
			if (String.IsNullOrWhiteSpace(sortExpression))
			{
				sortExpression = "DATECREATED desc";
			}
			//var userList = new List<int>();
			//if (userId > 0)
			//	userList.Add(userId);
			//else
			//	userList.AddRange(UserController.SelectUsers(postId).Select(user => user.ID));
			if (postIdFilter > 0)
			{
				postIdTree.Clear();
				postIdTree.Add(postIdFilter);
			}
			var bfStatusList = new List<int>();
			if (bfStatusId > 0)
				bfStatusList.Add(bfStatusId);
			else
			{
				bfStatusList.AddRange(BfStatusController.List().Select(x => x.ID));
				bfStatusList.Add(0); // значение по умолчанию
			}
			opsId = opsId ?? string.Empty;
			deviceModel = deviceModel ?? string.Empty;
			deviceModel = deviceModel.ToLower();
			bfNumber = bfNumber ?? string.Empty;
			bfNumber = bfNumber.ToLower().Trim();
			deviceNumber = deviceNumber ?? string.Empty;
			deviceNumber = deviceNumber.ToLower().Trim();
			parentBfNumber = parentBfNumber ?? string.Empty;
			parentBfNumber = parentBfNumber.ToLower().Trim();
			text = text ?? string.Empty;
			text = text.ToLower().Trim();
			using (hdEntities entities = new hdEntities())
			{
				List<REQUEST> r = new List<REQUEST>();
				if (maximumRows == 0 && startRowIndex == 0)
					r = (from x in entities.REQUEST
						 let post_id = x.POST_ID
						 let author_id = x.AUTHOR_ID
						 let bf_status_id = x.BF_STATUS_ID.HasValue ? x.BF_STATUS_ID.Value : 0
						 let solvedvalue = x.SOLVED.HasValue && x.SOLVED.Value
						 let dev0 = x.DEVICEINFOS.FirstOrDefault()
						 let datecreated = x.DATECREATED
						 let timezone = x.USER.POST.regions.timezone
						 let bf_number = x.BF_NUMBER
						 let parent_bf_number = x.parent_id.HasValue ? x.parent.BF_NUMBER : "-"
						 let _text = x.TEXT.ToLower().Trim()
						 where
							postIdTree.Contains(post_id) && (userId == 0 || userId != 0 && author_id == userId) &&
							bfStatusList.Contains(bf_status_id) &&
							(solved == 0 || solved == 1 && solvedvalue ||
							 solved == 2 && !solvedvalue) &&
							EntityFunctions.AddHours(datecreated, timezone) >= SqlDateTime.MinValue.Value &&
							EntityFunctions.AddHours(datecreated, timezone) <= SqlDateTime.MaxValue.Value &&
							SqlFunctions.StringConvert((double?)(dev0.OPS.idx)).Contains(opsId) &&
							(dev0.type ?? string.Empty).ToLower().Contains(deviceModel) &&
							(dev0.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) &&
							(bf_number ?? string.Empty).ToLower().Contains(bfNumber) &&
							parent_bf_number.ToLower().Contains(parentBfNumber) &&
							_text.Contains(text)
						 //&&
						 //(x.DEVICEINFO == null && (x.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) ||
						 //x.DEVICEINFO != null &&
						 //(x.DEVICEINFO.DEVICENUMBER ?? String.Empty).ToLower().Contains(deviceNumber))
						 select x).OrderBy(sortExpression).ToList();
				else
					r = (from x in entities.REQUEST
						 let post_id = x.POST_ID
						 let author_id = x.AUTHOR_ID
						 let bf_status_id = x.BF_STATUS_ID.HasValue ? x.BF_STATUS_ID.Value : 0
						 let solvedvalue = x.SOLVED.HasValue && x.SOLVED.Value
						 let dev0 = x.DEVICEINFOS.FirstOrDefault()
						 let datecreated = x.DATECREATED
						 let timezone = x.USER.POST.regions.timezone
						 let bf_number = x.BF_NUMBER
						 let parent_bf_number = x.parent_id.HasValue ? x.parent.BF_NUMBER : "-"
						 let _text = x.TEXT.ToLower().Trim()
						 where
							postIdTree.Contains(post_id) && (userId == 0 || userId != 0 && author_id == userId) &&
							bfStatusList.Contains(bf_status_id) &&
							(solved == 0 || solved == 1 && solvedvalue ||
							 solved == 2 && !solvedvalue) &&
							EntityFunctions.AddHours(datecreated, timezone) >= SqlDateTime.MinValue.Value &&
							EntityFunctions.AddHours(datecreated, timezone) <= SqlDateTime.MaxValue.Value &&
							SqlFunctions.StringConvert((double?)(dev0.OPS.idx)).Contains(opsId) &&
							(dev0.type ?? string.Empty).ToLower().Contains(deviceModel) &&
							(dev0.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) &&
							(bf_number ?? string.Empty).ToLower().Contains(bfNumber) &&
							parent_bf_number.ToLower().Contains(parentBfNumber) &&
							_text.Contains(text)
						 //&&
						 //(x.DEVICEINFO == null && (x.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) ||
						 //x.DEVICEINFO != null &&
						 //(x.DEVICEINFO.DEVICENUMBER ?? String.Empty).ToLower().Contains(deviceNumber))
						 select x).OrderBy(sortExpression).Skip(startRowIndex).Take(maximumRows).ToList();
				foreach (var request in r)
				{
					if (!request.POSTReference.IsLoaded)
						entities.LoadProperty(request, "POST");
					if (!request.BF_STATUSReference.IsLoaded)
						entities.LoadProperty(request, "BF_STATUS");
					if (!request.DEVICEINFOS.IsLoaded)
						entities.LoadProperty(request, "DEVICEINFOS");
					if (!request.DEVICEINFOS.First().OPSReference.IsLoaded)
						entities.LoadProperty(request.DEVICEINFOS.First(), x => x.OPS);
					if (!request.USERReference.IsLoaded)
						entities.LoadProperty(request, "USER");
					if (!request.USER.POSTReference.IsLoaded)
						entities.LoadProperty(request.USER, "POST");
					if (!request.USER.POST.regionsReference.IsLoaded)
						entities.LoadProperty(request.USER.POST, "regions");
					if (!request.servicesReference.IsLoaded)
						entities.LoadProperty(request, "services");

				}
				return r;
			}

		}

		public static REQUEST FindByInnerId(string requestInnerId)
		{
			using (hdEntities entities = new hdEntities())
			{
				REQUEST request = entities.REQUEST.SingleOrDefault(x => x.inner_number == requestInnerId);
				return request;
			}
		}


		private void searchPostIdTree(int postId)
		{
			using (hdEntities entities = new hdEntities())
			{
				if (!postIdTree.Contains(postId))
					postIdTree.Add(postId);
				var posts = from post in entities.POST where post.PARENT_ID == postId select post;
				foreach (POST post in posts)
					searchPostIdTree(post.ID);
			}
		}

		public int CountByPost(int postId, Int32 userId, int postIdFilter,
			string opsId, Int32 solved, string deviceModel, int bfStatusId, string bfNumber, string deviceNumber, string parentBfNumber, string text)
		{
			//return FindByPost(postId, null, userId, postIdFilter, SqlDateTime.MinValue.Value, SqlDateTime.MaxValue.Value,
			//	opsId, solved, deviceModel, bfStatusId, bfNumber, deviceNumber, parentBfNumber, text).Count();
			postIdTree = new List<int>();
			// postId прилетает от пользователя
			searchPostIdTree(postId);
			//var userList = new List<int>();
			//if (userId > 0)
			//	userList.Add(userId);
			//else
			//	userList.AddRange(UserController.SelectUsers(postId).Select(user => user.ID));
			if (postIdFilter > 0)
			{
				postIdTree.Clear();
				postIdTree.Add(postIdFilter);
			}
			var bfStatusList = new List<int>();
			if (bfStatusId > 0)
				bfStatusList.Add(bfStatusId);
			else
			{
				bfStatusList.AddRange(BfStatusController.List().Select(x => x.ID));
				bfStatusList.Add(0); // значение по умолчанию
			}
			opsId = opsId ?? string.Empty;
			deviceModel = deviceModel ?? string.Empty;
			deviceModel = deviceModel.ToLower();
			bfNumber = bfNumber ?? string.Empty;
			bfNumber = bfNumber.ToLower().Trim();
			deviceNumber = deviceNumber ?? string.Empty;
			deviceNumber = deviceNumber.ToLower().Trim();
			parentBfNumber = parentBfNumber ?? string.Empty;
			parentBfNumber = parentBfNumber.ToLower().Trim();
			text = text ?? string.Empty;
			text = text.ToLower().Trim();
			DateTime dateCreatedStart = SqlDateTime.MinValue.Value;
			DateTime dateCreatedEnd = SqlDateTime.MaxValue.Value;

			using (hdEntities entities = new hdEntities())
			{
				int r = 0;
				r = (from x in entities.REQUEST
					 let post_id = x.POST_ID
					 let author_id = x.AUTHOR_ID
					 let bf_status_id = x.BF_STATUS_ID.HasValue ? x.BF_STATUS_ID.Value : 0
					 let solvedvalue = x.SOLVED.HasValue && x.SOLVED.Value
					 let dev0 = x.DEVICEINFOS.FirstOrDefault()
					 let datecreated = x.DATECREATED
					 let timezone = x.USER.POST.regions.timezone
					 let bf_number = x.BF_NUMBER
					 let parent_bf_number = x.parent_id.HasValue ? x.parent.BF_NUMBER : "-"
					 let _text = x.TEXT.ToLower().Trim()
					 where
						postIdTree.Contains(post_id) && (userId == 0 || userId != 0 && author_id == userId) &&
						//x.DEVICEINFOS.Any(y => postList.Contains(y.OPS.idx_post.Value)) &&
						//dev0.OPS.idx_post == postIdFilter &&
						//flag &&
						bfStatusList.Contains(bf_status_id) &&
						(solved == 0 || solved == 1 && solvedvalue ||
						 solved == 2 && !solvedvalue) &&
						EntityFunctions.AddHours(datecreated, timezone) >= dateCreatedStart &&
						EntityFunctions.AddHours(datecreated, timezone) <= dateCreatedEnd &&
						SqlFunctions.StringConvert((double?)(dev0.OPS.idx)).Contains(opsId) &&
						(dev0.type ?? string.Empty).ToLower().Contains(deviceModel) &&
						(dev0.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) &&
						(bf_number ?? string.Empty).ToLower().Contains(bfNumber) &&
						parent_bf_number.ToLower().Contains(parentBfNumber) &&
						_text.Contains(text)

					 //&&
					 //(x.DEVICEINFO == null && (x.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) ||
					 //x.DEVICEINFO != null &&
					 //(x.DEVICEINFO.DEVICENUMBER ?? String.Empty).ToLower().Contains(deviceNumber))
					 select x).Count();
				return r;
			}
		}

		public int CountByPost(int postId, int userId, int postIdFilter, DateTime dateCreatedStart,
			DateTime dateCreatedEnd, string opsId, int solved, string deviceModel, int bfStatusId, string bfNumber,
			string deviceNumber, string parentBfNumber, string text)
		{

			postIdTree = new List<int>();
			// postId прилетает от пользователя
			searchPostIdTree(postId);
			//var userList = new List<int>();
			//if (userId > 0)
			//	userList.Add(userId);
			//else
			//	userList.AddRange(UserController.SelectUsers(postId).Select(user => user.ID));
			if (postIdFilter > 0)
			{
				postIdTree.Clear();
				postIdTree.Add(postIdFilter);
			}
			var bfStatusList = new List<int>();
			if (bfStatusId > 0)
				bfStatusList.Add(bfStatusId);
			else
			{
				bfStatusList.AddRange(BfStatusController.List().Select(x => x.ID));
				bfStatusList.Add(0); // значение по умолчанию
			}
			opsId = opsId ?? string.Empty;
			deviceModel = deviceModel ?? string.Empty;
			deviceModel = deviceModel.ToLower();
			bfNumber = bfNumber ?? string.Empty;
			bfNumber = bfNumber.ToLower().Trim();
			deviceNumber = deviceNumber ?? string.Empty;
			deviceNumber = deviceNumber.ToLower().Trim();
			parentBfNumber = parentBfNumber ?? string.Empty;
			parentBfNumber = parentBfNumber.ToLower().Trim();
			text = text ?? string.Empty;
			text = text.ToLower().Trim();

			using (hdEntities entities = new hdEntities())
			{
				int r = 0;
				r = (from x in entities.REQUEST
					 let post_id = x.POST_ID
					 let author_id = x.AUTHOR_ID
					 let bf_status_id = x.BF_STATUS_ID.HasValue ? x.BF_STATUS_ID.Value : 0
					 let solvedvalue = x.SOLVED.HasValue && x.SOLVED.Value
					 let dev0 = x.DEVICEINFOS.FirstOrDefault()
					 let datecreated = x.DATECREATED
					 let timezone = x.USER.POST.regions.timezone
					 let bf_number = x.BF_NUMBER
					 let parent_bf_number = x.parent_id.HasValue ? x.parent.BF_NUMBER : "-"
					 let _text = x.TEXT.ToLower().Trim()
					 where
						postIdTree.Contains(post_id) && (userId == 0 || userId != 0 && author_id == userId) &&
						//x.DEVICEINFOS.Any(y => postList.Contains(y.OPS.idx_post.Value)) &&
						//dev0.OPS.idx_post == postIdFilter &&
						//flag &&
						bfStatusList.Contains(bf_status_id) &&
						(solved == 0 || solved == 1 && solvedvalue ||
						 solved == 2 && !solvedvalue) &&
						EntityFunctions.AddHours(datecreated, timezone) >= dateCreatedStart &&
						EntityFunctions.AddHours(datecreated, timezone) <= dateCreatedEnd &&
						SqlFunctions.StringConvert((double?)(dev0.OPS.idx)).Contains(opsId) &&
						(dev0.type ?? string.Empty).ToLower().Contains(deviceModel) &&
						(dev0.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) &&
						(bf_number ?? string.Empty).ToLower().Contains(bfNumber) &&
						parent_bf_number.ToLower().Contains(parentBfNumber) &&
						_text.Contains(text)

					 //&&
					 //(x.DEVICEINFO == null && (x.DEVICENUMBER ?? string.Empty).ToLower().Contains(deviceNumber) ||
					 //x.DEVICEINFO != null &&
					 //(x.DEVICEINFO.DEVICENUMBER ?? String.Empty).ToLower().Contains(deviceNumber))
					 select x).Count();
				return r;
			}
		}

		//public IEnumerable<REQUEST> FindByPost(int postId)
		//{
		//    var requests = from r in entities.REQUEST where r.POST_ID == postId select r;
		//    foreach (var request in requests)
		//    {
		//        yield return request;
		//    }
		//    var posts = from post in entities.POST where post.PARENT_ID == postId select post;
		//    foreach (var post in posts)
		//    {
		//        var recursiveSequence = FindByPost(post.ID);
		//        foreach (var request in recursiveSequence)
		//        {
		//            yield return request;
		//        }
		//    }
		//}

		private static void setCell(XWPFTable table, int row, int col, string content, bool bold = false, int size = 14)
		{
			var paraRun = table.Rows[row].GetCell(col).Paragraphs[0].CreateRun();
			paraRun.SetText(content);
			paraRun.FontSize = size;
			paraRun.SetBold(bold);
		}

		public static Stream CreateDocFile(EmailConfigSection config, REQUEST request, ref string fileName)
		{
			using (hdEntities entities = new hdEntities())
			{
				entities.Attach(request);
				////entities.LoadProperty(request.POST, x => x.regions);
				//entities.Entry(request.POST).Reference(x => x.regions).Load();
				////entities.LoadProperty(request, x => x.OPS);
				//entities.Entry(request).Reference(x => x.OPS).Load();
				////entities.LoadProperty(request, x => x.USER);
				//entities.Entry(request).Reference(x => x.USER).Load();
				////entities.LoadProperty(request.USER, x => x.POST);
				//entities.Entry(request.USER).Reference(x => x.POST).Load();
				////entities.LoadProperty(request.USER.POST, x => x.regions);
				//entities.Entry(request.USER.POST).Reference(x => x.regions).Load();
				////entities.LoadProperty(request, x => x.DEVICEINFOS);
				//entities.Entry(request).Collection(x => x.DEVICEINFOS).Load();
				entities.LoadProperty(request.POST, x => x.regions);
				entities.LoadProperty(request, x => x.OPS);
				entities.LoadProperty(request, x => x.USER);
				entities.LoadProperty(request.USER, x => x.POST);
				entities.LoadProperty(request.USER.POST, x => x.regions);
				entities.LoadProperty(request, x => x.DEVICEINFOS);

				//if (!request.MULTIPLE.HasValue || request.MULTIPLE.HasValue && request.MULTIPLE.Value == false)
				{
					XWPFDocument docx = new XWPFDocument();
					var para = docx.CreateParagraph();
					var paraRun = para.CreateRun();
					paraRun.SetText(string.Format("АДРЕС ОТПРАВКИ: {0}", config.Author));
					paraRun.FontFamily = "Calibri";
					paraRun.FontSize = 14;

					int rowCount = 9, colCount = 3;

					var table = docx.CreateTable(rowCount, colCount);
					setCell(table, 0, 0, "№", true);
					setCell(table, 0, 1, "Информационные поля заявки", true);
					setCell(table, 0, 2, "Поля для заполнения", true);

					for (int i = 1; i < rowCount; i++)
					{
						setCell(table, i, 0, i.ToString(), true);
					}
					setCell(table, 1, 1, "Заказчик");
					setCell(table, 2, 1, "Головное отделение (почтамт)");
					setCell(table, 3, 1, "Индекс и адрес Объекта обслуживания (ОПС, подразделение)");
					//setCell(table, 4, 1, "Winpost и ПО(восстановление работоспособности ПО)");
					//setCell(table, 5, 1, "Winpost и ПО (обновление и установка нового ПО)");
					setCell(table, 4, 1, "Наименование оборудования");
					setCell(table, 5, 1, "Идентификатор оборудования");
					setCell(table, 6, 1, "Описание заявки");
					//setCell(table, 9, 1, "Описание неисправности");
					setCell(table, 7, 1, "Инициатор, заявитель (ФИО, должность, № телефона)");
					setCell(table, 8, 1, "Контактное лицо на Объекте и № телефона");
					//setCell(table, 12, 1, "Вложение (документы, фото)");

					setCell(table, 1, 2, request.USER.POST.regions.name);
					setCell(table, 2, 2, "См. «Адресную базу»");
					setCell(table, 3, 2, "См. «Адресную базу»");
					//setCell(table, 4, 2, request.WINPOST_REPAIR.HasValue ? request.WINPOST_REPAIR.Value ? "Да" : "Нет" : "Нет");
					//setCell(table, 5, 2, request.WINPOST_UPDATE.HasValue ? request.WINPOST_UPDATE.Value ? "Да" : "Нет" : "Нет");
					setCell(table, 4, 2, "См. «Адресную базу»");
					setCell(table, 5, 2, "См. «Адресную базу»");
					setCell(table, 6, 2, request.TEXT);
					setCell(table, 7, 2, request.CONTACT);
					setCell(table, 8, 2, request.CONTACT);
					//string fileName;
					do
					{
						System.Threading.Thread.Sleep(10);
						fileName = Path.Combine(Path.GetTempPath(),
							string.Format("{0:yyyyMMddHHmmssfff}.doc", DateTime.Now));
						//System.Threading.Thread.Sleep(10);
					} while (File.Exists(fileName));
					FileStream fs = new FileStream(fileName, FileMode.Create);
					docx.Write(fs);
					return fs;
				}

				/*else
					{
						XWPFDocument docx = new XWPFDocument();
						var para = docx.CreateParagraph();
						var paraRun = para.CreateRun();
						paraRun.SetText(string.Format("АДРЕС ОТПРАВКИ: {0}", config.Author));
						paraRun.FontFamily = "Calibri";
						paraRun.FontSize = 14;

						int rowCount = 9, colCount = 3;

						var table = docx.CreateTable(rowCount, colCount);
						setCell(table, 0, 0, "№", true);
						setCell(table, 0, 1, "Информационные поля заявки", true);
						setCell(table, 0, 2, "Поля для заполнения", true);

						for (int i = 1; i < rowCount; i++)
						{
							setCell(table, i, 0, i.ToString(), true);
						}
						setCell(table, 1, 1, "Юридическое лицо");
						setCell(table, 2, 1, "Почтамт (головное отделение)");
						setCell(table, 3, 1, "Индекс ОПС и адрес (отделение)");
						setCell(table, 4, 1, "Winpost и ПО(восстановление работоспособности ПО)");
						setCell(table, 5, 1, "Winpost и ПО (обновление и установка нового ПО)");
						setCell(table, 6, 1, "Модель");
						setCell(table, 7, 1, "Серийный номер");
						setCell(table, 8, 1, "Предмет заявки (суть обращения)");
						setCell(table, 9, 1, "Описание неисправности");
						setCell(table, 10, 1, "Контактное лицо и № телефона");
						setCell(table, 11, 1, "Инициатор, заявитель (ФИО, должность, № телефона)");
						setCell(table, 12, 1, "Вложение (документы, фото)");

						setCell(table, 1, 2, "УФПС Алтайского края - филиал ФГУП «Почта России»");
						setCell(table, 2, 2, request.POST.NAME);
						StringBuilder opsAddresses = new StringBuilder();
						foreach (var ops in request.OPSES)
						{
							string opsAddress = string.Format("{0}, {1}, {2} {3}, {4}{5}", ops.region, ops.name_ops,
															  ops.street_type, ops.street,
															  ops.house, ops.litera);
							if (ops.corpus != null && ops.corpus > 0)
								opsAddress += " корп. " + ops.corpus;
							opsAddresses.AppendLine(String.Format("{1} {0};\r\n", opsAddress, ops.idx));

						}

						setCell(table, 3, 2, opsAddresses.ToString());
						setCell(table, 4, 2, request.WINPOST_REPAIR.HasValue ? request.WINPOST_REPAIR.Value ? "Да" : "Нет" : "Нет");
						setCell(table, 5, 2, request.WINPOST_UPDATE.HasValue ? request.WINPOST_UPDATE.Value ? "Да" : "Нет" : "Нет");
						setCell(table, 6, 2, request.DEVICEMODEL);
						setCell(table, 7, 2, request.DEVICENUMBER);
						setCell(table, 8, 2, request.SHORTTEXT);
						setCell(table, 9, 2, request.TEXT);
						string opsPhones = "-";
						opsPhones = string.Empty;
						setCell(table, 10, 2, opsPhones);
						setCell(table, 11, 2, request.CONTACT);
						setCell(table, 12, 2, "-");
						string fileName;
						do
						{
							fileName = Path.Combine(Path.GetTempPath(),
													string.Format("{0:yyyyMMddHHmmssfff}.doc", DateTime.Now));
							System.Threading.Thread.Sleep(10);
						} while (File.Exists(fileName));
						FileStream fs = new FileStream(fileName, FileMode.Create);
						docx.Write(fs);
						return fs;
					}
				 */
			}
		}

		public static Stream CreateXlsFile(REQUEST request, string fileName, string templatesDirectory)
		{
			//if (!request.MULTIPLE.HasValue || request.MULTIPLE.HasValue && request.MULTIPLE.Value == false)
			{
				using (hdEntities entities = new hdEntities())
				{
					entities.Attach(request);
					fileName = Path.ChangeExtension(fileName, ".xls");
					using (
						FileStream fs = new FileStream(Path.Combine(templatesDirectory, "template.xls"), FileMode.Open))
					{
						HSSFWorkbook workbook = new HSSFWorkbook(fs);
						HSSFSheet sheet = (HSSFSheet)workbook.GetSheetAt(0);
						int startRow = sheet.FirstRowNum + 10;
						for (int i = 0; i < request.DEVICEINFOS.Count - 1; i++)
							sheet.CopyRow(startRow, startRow + i + 1);
						for (int i = 0; i < request.DEVICEINFOS.Count; i++)
						{
							var row = sheet.GetRow(startRow + i);
							entities.LoadProperty(request.DEVICEINFOS.ElementAt(i), x => x.OPS);
							//entities.Entry(request.DEVICEINFOS.ElementAt(i)).Reference(x => x.OPS).Load();
							entities.LoadProperty(request.DEVICEINFOS.ElementAt(i).OPS, x => x.POST);
							//entities.Entry(request.DEVICEINFOS.ElementAt(i).OPS).Reference(x => x.POST).Load();
							row.Cells[row.FirstCellNum].SetCellValue(request.DEVICEINFOS.ElementAt(i).OPS.location);
							row.Cells[row.FirstCellNum + 1].SetCellValue(request.DEVICEINFOS.ElementAt(i).OPS.idx);
							row.Cells[row.FirstCellNum + 2].SetCellValue(request.DEVICEINFOS.ElementAt(i).OPS.POST.NAME);
							row.Cells[row.FirstCellNum + 3].SetCellValue(request.DEVICEINFOS.ElementAt(i).type);
							row.Cells[row.FirstCellNum + 4].SetCellValue(request.DEVICEINFOS.ElementAt(i).DEVICENUMBER);
							row.Cells[row.FirstCellNum + 5].SetCellValue(request.DEVICEINFOS.ElementAt(i).mfr);
							row.Cells[row.FirstCellNum + 6].SetCellValue(request.DEVICEINFOS.ElementAt(i).DEVICEMODEL);
						}

						//string fileName;
						using (FileStream fs1 = new FileStream(fileName, FileMode.Create))
						{
							workbook.Write(fs1);
							return fs1;
						}
					}
				}
			}
		}

		public static void SendEmail_new(object req, string templatesDirectory)
		{


			EmailConfigSection config =
				(EmailConfigSection)ConfigurationManager.GetSection(

					"EmailConfigGroup/emailConfig");
			REQUEST request = (REQUEST)req;
			if (request.MULTIPLE.HasValue && !request.MULTIPLE.Value)
			{
				string fileName1 = "";
				var docFileStream1 = CreateDocFile1(config, request, ref fileName1);
				SendEmailDoc(request, config, docFileStream1);
				return;
			}
			string fileName = "";
			var docFileStream = CreateDocFile(config, request, ref fileName);
			var xlsFileStream = CreateXlsFile(request, fileName, templatesDirectory);
			SendEmailDoc(request, config, docFileStream, xlsFileStream);
			return;
		}

		private static Stream CreateDocFile1(EmailConfigSection config, REQUEST request, ref string fileName)
		{
			using (hdEntities entities = new hdEntities())
			{
				entities.Attach(request);
				////entities.LoadProperty(request.POST, x => x.regions);
				//entities.Entry(request.POST).Reference(x => x.regions).Load();
				////entities.LoadProperty(request, x => x.OPS);
				//entities.Entry(request).Reference(x => x.OPS).Load();
				////entities.LoadProperty(request, x => x.USER);
				//entities.Entry(request).Reference(x => x.USER).Load();
				////entities.LoadProperty(request.USER, x => x.POST);
				//entities.Entry(request.USER).Reference(x => x.POST).Load();
				////entities.LoadProperty(request.USER.POST, x => x.regions);
				//entities.Entry(request.USER.POST).Reference(x => x.regions).Load();
				////entities.LoadProperty(request, x => x.DEVICEINFOS);
				//entities.Entry(request).Collection(x => x.DEVICEINFOS).Load();
				entities.LoadProperty(request.POST, x => x.regions);
				entities.LoadProperty(request, x => x.OPS);
				entities.LoadProperty(request, x => x.USER);
				entities.LoadProperty(request.USER, x => x.POST);
				entities.LoadProperty(request.USER.POST, x => x.regions);
				entities.LoadProperty(request, x => x.DEVICEINFOS);


				//if (!request.MULTIPLE.HasValue || request.MULTIPLE.HasValue && request.MULTIPLE.Value == false)
				{
					XWPFDocument docx = new XWPFDocument();
					var para = docx.CreateParagraph();
					var paraRun = para.CreateRun();
					paraRun.SetText(string.Format("АДРЕС ОТПРАВКИ: {0}", config.Author));
					paraRun.FontFamily = "Calibri";
					paraRun.FontSize = 14;

					int rowCount = 9, colCount = 3;

					var table = docx.CreateTable(rowCount, colCount);
					setCell(table, 0, 0, "№", true);
					setCell(table, 0, 1, "Информационные поля заявки", true);
					setCell(table, 0, 2, "Поля для заполнения", true);

					for (int i = 1; i < rowCount; i++)
					{
						setCell(table, i, 0, i.ToString(), true);
					}
					setCell(table, 1, 1, "Заказчик");
					setCell(table, 2, 1, "Головное отделение (почтамт)");
					setCell(table, 3, 1, "Индекс и адрес Объекта обслуживания (ОПС, подразделение)");
					//setCell(table, 4, 1, "Winpost и ПО(восстановление работоспособности ПО)");
					//setCell(table, 5, 1, "Winpost и ПО (обновление и установка нового ПО)");
					setCell(table, 4, 1, "Наименование оборудования");
					setCell(table, 5, 1, "Идентификатор оборудования");
					setCell(table, 6, 1, "Описание заявки");
					//setCell(table, 9, 1, "Описание неисправности");
					setCell(table, 7, 1, "Инициатор, заявитель (ФИО, должность, № телефона)");
					setCell(table, 8, 1, "Контактное лицо на Объекте и № телефона");
					//setCell(table, 12, 1, "Вложение (документы, фото)");
					entities.LoadProperty(request.DEVICEINFOS.ElementAt(0), x => x.OPS);
					//entities.Entry(request.DEVICEINFOS.ElementAt(0)).Reference(x => x.OPS).Load();
					entities.LoadProperty(request.DEVICEINFOS.ElementAt(0).OPS, x => x.POST);
					//entities.Entry(request.DEVICEINFOS.ElementAt(0).OPS).Reference(x => x.POST).Load();
					setCell(table, 1, 2, request.USER.POST.regions.name);
					setCell(table, 2, 2, request.DEVICEINFOS.ElementAt(0).OPS.POST.NAME);
					setCell(table, 3, 2, request.DEVICEINFOS.ElementAt(0).OPS.idx + " " + request.DEVICEINFOS.ElementAt(0).OPS.location);
					setCell(table, 4, 2, request.DEVICEINFOS.ElementAt(0).type + "," + request.DEVICEINFOS.ElementAt(0).DEVICEMODEL);
					setCell(table, 5, 2, request.DEVICEINFOS.ElementAt(0).DEVICENUMBER);
					setCell(table, 6, 2, request.TEXT);
					setCell(table, 7, 2, request.CONTACT);
					string opsContact = request.DEVICEINFOS.ElementAt(0).OPS.phone;
					if (string.IsNullOrEmpty(opsContact))
						opsContact = request.CONTACT;
					setCell(table, 8, 2, opsContact);
					//string fileName;
					do
					{
						System.Threading.Thread.Sleep(10);
						fileName = Path.Combine(Path.GetTempPath(),
							string.Format("{0:yyyyMMddHHmmssfff}.doc", DateTime.Now));
						//System.Threading.Thread.Sleep(10);
					} while (File.Exists(fileName));
					FileStream fs = new FileStream(fileName, FileMode.Create);
					docx.Write(fs);
					return fs;
				}
			}
		}

		private static void SendEmailDoc(REQUEST request, EmailConfigSection config, Stream docFileStream,
			Stream xlsFileStream = null)
		{
			string emails = config.Email;
			ParameterizedThreadStart start = SendMessage;
			FileStream fs = (FileStream)docFileStream;
			FileStream fs1 = xlsFileStream as FileStream;
			List<Thread> threads = new List<Thread>();

			ExchangeService exchangeService = ExchangeServiceCreator.Create();

			foreach (var email in emails.Split(','))
			{
				string requestSubjectValid = request.SHORTTEXT.Replace('\r', ' ');
				requestSubjectValid = requestSubjectValid.Replace('\n', ' ');
				var mailMessage = new EmailMessage(exchangeService);
				mailMessage.ToRecipients.Add(email);
				mailMessage.From = config.Author;
				mailMessage.Subject = string.Format("[{0}] {1}", request.inner_number, requestSubjectValid);
				var docFileContents = File.ReadAllBytes(fs.Name);
				mailMessage.Attachments.AddFileAttachment("request.docx", docFileContents);
				if (fs1 != null)
				{
					var xlsFileContents = File.ReadAllBytes(fs1.Name);
					mailMessage.Attachments.AddFileAttachment("адресная база.xls", xlsFileContents);
				}
				mailMessage.Save(WellKnownFolderName.SentItems);
				SaveEmailAsFileToRequestAndDisk(request, request.DATECREATED, mailMessage, mailMessage.Subject, mailMessage.Id.ToString());
				//mailMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.Delay | DeliveryNotificationOptions.OnFailure | DeliveryNotificationOptions.OnSuccess;
				//SendMessage(mailMessage);
				var t = new Thread(start);
				t.Start(mailMessage);
			}
			//File.Delete(fs.Name);
		}

		private static bool checkForThreadsToFinish(List<Thread> threads)
		{
			bool res = true;
			foreach (Thread thread in threads)
			{
				res = res && thread.ThreadState == ThreadState.Stopped;
			}
			return res;
		}

		public static void SendMessage(object arg)
		{

			string msg = "no data";

			EmailMessage mailmsg = (EmailMessage)arg;
			mailmsg.Load();
			msg = mailmsg.From.ToString() + "|=|" + mailmsg.ToRecipients.ToString();
			bool f = false;
			while (!f)
			{
				try
				{
					mailmsg.SendAndSaveCopy();
					f = true;
				}
				catch (Exception e)
				{
					WebLogger.GetInstance().Error("error={0};msg={1};", e, msg);
					Thread.Sleep(2000);
				}

			}
			GC.Collect();
		}

		public static REQUEST FindById(int id)
		{
			using (hdEntities entities = new hdEntities())
			{
				REQUEST r =
					entities.REQUEST.Include("POST")
						.Include("USER.POST.regions")
						.Include("BF_STATUS")
						.Include("DEVICEINFOS")
						.Include("DEVICEINFOS.OPS")
						.Include("services")
						.SingleOrDefault(x => x.ID == id);
				return r;
			}
		}

		public static REQUEST FindById(hdEntities entities, int id)
		{
			REQUEST r =
				entities.REQUEST.Include("POST")
					.Include("USER.POST.regions")
					.Include("BF_STATUS")
					.Include("DEVICEINFOS")
					.Include("services")
					.SingleOrDefault(x => x.ID == id);
			return r;
		}

		public static bool isRequestSolved(int requestId)
		{
			REQUEST request = FindById(requestId);
			if (request == null)
				return true;
			bool isSolved = request.SOLVED.HasValue;
			if (isSolved)
				isSolved = isSolved && request.SOLVED.Value;
			return isSolved;
		}

		public static void SolveRequest(int requestId, string actNumber, DateTime solvedDateTime, USER solverUser)
		{
			using (hdEntities entities = new hdEntities())
			{
				REQUEST request = FindById(entities, requestId);
				if (request == null)
					return;
				request.NUMBER = actNumber;
				request.DATESOLVED = solvedDateTime.AddHours(-1 * request.USER.POST.regions.timezone);
				request.act_date = request.DATESOLVED;
				request.SOLVED = true;
				request.SOLVER_ID = solverUser.ID;
				BF_STATUS bfStatus = entities.BF_STATUS.Where(x => x.TEXT.ToLower().Contains("закр")).Single();
				request.BF_STATUS = bfStatus;
				Update(entities, request);
			}
		}

		public static normatives SelectNormative(REQUEST request, out string message, List<normatives> normativesList)
		{
			if (request.MULTIPLE.HasValue && request.MULTIPLE.Value)
			{
				message = "";
				return null;
			}
			var norm = NormativeController.Select(request.service_id,
				request.DEVICEINFOS.First().OPS.type_city ?? "1",
				request.DEVICEINFOS.First().OPS.priority ?? "0",
				request.DEVICEINFOS.First().devicetype_id, out message, request.DATECREATED, normativesList);
			return norm;
		}

		public static List<periods> GetPeriodsForDate(DateTime currentDate, List<periods> periods, List<days_off> daysOffList)
		{
			List<periods> resPeriods = null;
			int tekDay = ((int)currentDate.DayOfWeek + 6) % 7;
			// выбор из праздничных дней текущей даты
			days_off dayOff = daysOffList.Where(x => x.date.Date == currentDate.Date).SingleOrDefault();
			bool[] workDays = new bool[7];
			foreach (var period in periods)
				workDays[period.day_number.Value] = true;
			//флаг для итогового выходного
			//если нет праздников
			if (dayOff == null)
			{
				if (workDays[tekDay])
				{
					resPeriods = periods.Where(x => x.day_number == tekDay).ToList();
				}
			}
			else
			{
				// если праздник не выходной
				if (!dayOff.day_off)
				{
					//получить из базы реальный номер дня
					int daynumber = tekDay;
					if (dayOff.day_correction >= 0)
						daynumber = dayOff.day_correction;
					if (workDays[daynumber])
					{
						resPeriods = new List<periods>();
						foreach (var p in periods.Where(x => x.day_number == daynumber).ToList())
						{
							resPeriods.Add(new periods()
							{
								start = p.start,
								finish = p.finish,
								number = p.number,
								day_number = p.day_number
							});
						}
						if (dayOff.time_correction != 0)
						{
							periods lastDayPeriod = resPeriods.Last();
							TimeSpan newLastTime = lastDayPeriod.finish.Value.TimeOfDay;
							newLastTime = newLastTime.Add(TimeSpan.FromHours(dayOff.time_correction));
							resPeriods[resPeriods.Count - 1].finish = new DateTime(lastDayPeriod.finish.Value.Year,
								lastDayPeriod.finish.Value.Month,
								lastDayPeriod.finish.Value.Day, newLastTime.Hours, newLastTime.Minutes, newLastTime.Seconds);
						}
					}
				}
			}
			return resPeriods;
		}

		public static DateTime SelectNormativeDate(REQUEST request, normatives norma, out string message, IList<periods> periodList = null)
		{
			//	OpsController.FindById(request.DEVICEINFOS.ElementAt(0).OPS_ID.Value).periods.Load();
			List<periods> periods;
			int ops_id = request.DEVICEINFOS.First().OPS_ID.Value;
			if (periodList == null)
			{
				periods = OpsController.FindById(ops_id).periods.OrderBy(x => x.number).ToList();
				//if (!request.DEVICEINFOS.First().OPS.periods.IsLoaded)
				//	OpsController.LoadPeriods(request.DEVICEINFOS.First().OPS);
				//periods = request.DEVICEINFOS.First().OPS.periods.ToList();
			}
			else
			{

				periods = periodList.Where(x => x.ops_id == ops_id).OrderBy(x => x.number).ToList();
			}
			message = "";
			if (periods.Count == 0)
			{
				message = " Для данного ОПС не указан график работы";
				return DateTime.MinValue;
			}
			List<days_off> daysOffList = NormativeController.DaysOffList();
			// развертка списка графиков работ на 80 дней вперед для упрощения работы с выходными днями
			List<periods> periodsLarge = new List<periods>();
			List<DateTime> datesLarge = new List<DateTime>();
			// прибавление разницы в часовых поясах
			DateTime requestDate = request.DATECREATED.AddHours(request.USER.POST.regions.timezone);
			DateTime currentDate = requestDate;
			List<periods> tempPeriods;
			// множество, для хранения дат рабочих дней
			HashSet<DateTime> workingDaysSet = new HashSet<DateTime>();
			do
			{
				currentDate = currentDate.AddDays(-1);
				tempPeriods = GetPeriodsForDate(currentDate, periods, daysOffList);
			} while (tempPeriods == null);
			foreach (var tempPeriod in tempPeriods)
			{
				periodsLarge.Add(tempPeriod);
				datesLarge.Add(currentDate);
				workingDaysSet.Add(currentDate.Date);
			}
			int counter = 0;
			for (DateTime iDate = requestDate; counter < 80; iDate = iDate.AddDays(1))
			{
				tempPeriods = GetPeriodsForDate(iDate, periods, daysOffList);
				if (tempPeriods != null)
					foreach (var tempPeriod in tempPeriods)
					{
						periodsLarge.Add(tempPeriod);
						datesLarge.Add(iDate);
						workingDaysSet.Add(iDate.Date);
					}
				counter++;
			}

			// во-первых, определить в рабочее ли время подана заявка
			int dayNumber = ((int)requestDate.DayOfWeek + 6) % 7;
			int startDay = 0;
			int startPeriodIndex = 0;
			// признак подачи заявки в рабочее время
			bool inWorkTime = false;
			// признак изменения планового времени заявки на окончание рабочего дня
			bool changeNormativeTimeToFinish = false;
			DateTime normativeDate = requestDate;
			// если заявка подана в рабочее время, то началом отсчета является время подачи заявки
			for (int i = 0; i < periodsLarge.Count; i++)
				if (requestDate.TimeOfDay.CompareTo(periodsLarge[i].start.Value.TimeOfDay) >= 0 &&
					requestDate.TimeOfDay.CompareTo(periodsLarge[i].finish.Value.TimeOfDay) <= 0 &&
					datesLarge[i].Date == requestDate.Date)
				{
					startDay = periods[i].day_number.Value;
					startPeriodIndex = i;
					inWorkTime = true;
					changeNormativeTimeToFinish = true;
					break;
				}
			// если заявка подана в нерабочее время, то находится ближайший в будущем рабочий период 
			// и от его начала начинается отсчет
			if (!inWorkTime)
			{
				for (int i = 0; i < periodsLarge.Count - 1; i++)
				{
					DateTime currentPeriodDate = new DateTime(datesLarge[i].Year, datesLarge[i].Month, datesLarge[i].Day,
						periodsLarge[i].finish.Value.Hour, periodsLarge[i].finish.Value.Minute, periodsLarge[i].finish.Value.Second);
					periods nextPeriod = periodsLarge[i + 1];
					DateTime nextPeriodDate = new DateTime(datesLarge[i + 1].Year, datesLarge[i + 1].Month, datesLarge[i + 1].Day,
						nextPeriod.start.Value.Hour, nextPeriod.start.Value.Minute, nextPeriod.start.Value.Second);
					if (requestDate.CompareTo(currentPeriodDate) > 0
						&& requestDate.CompareTo(nextPeriodDate) < 0)
					{
						startDay = nextPeriod.day_number.Value;
						startPeriodIndex = i + 1;
						normativeDate = nextPeriodDate;
						break;
					}
				}
			}
			// если дата отсчета равна дате подачи заявки,
			// то признак изменения планового времени заявки на окончание рабочего дня равен истине
			if (requestDate.Date == normativeDate.Date)
				changeNormativeTimeToFinish = true;


			// если срок в днях, то расчет проще
			// если день выходной, к дате прибавить день
			// если рабочий, к дате прибавить день, от срока отнять день
			// повторять, пока срок болше нуля
			if (norma.dayFlag.HasValue && norma.dayFlag.Value)
			{
				int daysLeft = norma.time;
				int daysCount = 40;
				while (daysLeft > 0 && daysCount > 0)
					if (workingDaysSet.Contains(normativeDate.Date))
					{
						normativeDate = normativeDate.AddDays(1);
						daysLeft--;
						daysCount--;
					}
					else
					{
						normativeDate = normativeDate.AddDays(1);
						daysCount--;
					}
				//пропуск выходных дней
				while (!workingDaysSet.Contains(normativeDate.Date) && daysCount > 0)
				{
					normativeDate = normativeDate.AddDays(1);
					daysCount--;
				}
			}
			else
			{
				bool stopFlag = false;
				double timeLeft = norma.time;
				// если время в часах
				if (startPeriodIndex < periods.Count)
					// повторять пока оставшееся время будет больше рабочего времени текущего периода
					for (int i = startPeriodIndex; !stopFlag; i++)
					{
						if (timeLeft > (periodsLarge[i].finish.Value.TimeOfDay - normativeDate.TimeOfDay).TotalHours)
						{
							// вычесть разницу между текущим началом и концом рабочего периода
							timeLeft -= (periodsLarge[i].finish.Value.TimeOfDay - normativeDate.TimeOfDay).TotalHours;
							periods nextPeriod = periodsLarge[i + 1];
							normativeDate = new DateTime(datesLarge[i + 1].Year, datesLarge[i + 1].Month, datesLarge[i + 1].Day,
								nextPeriod.start.Value.Hour,
								nextPeriod.start.Value.Minute, nextPeriod.start.Value.Second);
						}
						else
						{
							normativeDate = normativeDate.AddHours(timeLeft);
							stopFlag = true;
						}
					}
			}
			// в связи с новыми требованиями ЦРТ, введено ограничение на срок выполнения в календарных днях
			if (norma.dayFlag.Value && norma.time <= 15)
			{
				DateTime extendedRequestDate = requestDate.AddDays(21);
				if (extendedRequestDate < normativeDate)
				{
					normativeDate = extendedRequestDate;
					message = " Целевая дата ограничена 21 календарным днем";
				}
			}
			else if (norma.dayFlag.Value && norma.time <= 30)
			{
				DateTime extendedRequestDate = requestDate.AddDays(30);
				if (extendedRequestDate < normativeDate)
				{
					normativeDate = extendedRequestDate;
					message = " Целевая дата ограничена 30 календарными днями";
				}
			}
			//if (norma.dayFlag.Value)
			//	if (changeNormativeTimeToFinish)
			//	{
			//		int index = -1;
			//		for (int i = datesLarge.Count - 1; i >= 0 && index < 0; i--)
			//			if (datesLarge[i].Date == normativeDate.Date)
			//				index = i;
			//		if (index >= 0)
			//			normativeDate = new DateTime(normativeDate.Year, normativeDate.Month, normativeDate.Day,
			//				periodsLarge[index].finish.Value.Hour, periodsLarge[index].finish.Value.Minute, periodsLarge[index].finish.Value.Second);
			//	}

			return normativeDate;
		}

		public static void CreateReport(HttpContext context, IEnumerable requests, IDictionary parametersValues)
		{

			HSSFWorkbook hssfworkbook = new HSSFWorkbook();

			//create a entry of DocumentSummaryInformation
			DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
			dsi.Company = "ALTPOST.RU";
			hssfworkbook.DocumentSummaryInformation = dsi;

			////create a entry of SummaryInformation
			SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
			si.Subject = "REQUESTS REPORT";
			hssfworkbook.SummaryInformation = si;
			ISheet sheet1 = hssfworkbook.CreateSheet("Отчет");

			int i = 0;
			IRow ro = sheet1.CreateRow(i);
			ro.CreateCell(0).SetCellValue("Автор");
			ro.CreateCell(1).SetCellValue("Почтамт");
			ro.CreateCell(2).SetCellValue("Дата создания");
			ro.CreateCell(3).SetCellValue("Индекс ОПС");
			ro.CreateCell(4).SetCellValue("Класс");
			ro.CreateCell(5).SetCellValue("Наименование и модель оборудования");
			ro.CreateCell(6).SetCellValue("Серийный № оборудования");
			ro.CreateCell(7).SetCellValue("Вид услуги");
			ro.CreateCell(8).SetCellValue("Содержание");
			ro.CreateCell(9).SetCellValue("Заявка закрыта");

			ro.CreateCell(10).SetCellValue("№ Заявки ЦРТ");
			ro.CreateCell(11).SetCellValue("Статус ЦРТ");
			ro.CreateCell(12).SetCellValue("Примечание ЦРТ");

			ro.CreateCell(13).SetCellValue("Дата закрытия");
			ro.CreateCell(14).SetCellValue("Номер акта");
			ro.CreateCell(15).SetCellValue("Время работы");
			if ((int)parametersValues["postId"] == 656700)
			{
				ro.CreateCell(16).SetCellValue("Нормативное время");
				ro.CreateCell(17).SetCellValue("Просроченность");
			}
			IFont fontBold = hssfworkbook.CreateFont();
			fontBold.Boldweight = (short)FontBoldWeight.Bold;
			fontBold.FontHeightInPoints = 10;
			fontBold.FontName = "Arial";
			fontBold.IsItalic = false;
			//ro.RowStyle=new HSSFCellStyle(0,new ExtendedFormatRecord(),hssfworkbook );
			//ro.RowStyle.SetFont(fontBold);
			var boldCellStyle = hssfworkbook.CreateCellStyle();
			boldCellStyle.SetFont(fontBold);
			boldCellStyle.WrapText = true;
			boldCellStyle.BorderRight =
				boldCellStyle.BorderTop =
					boldCellStyle.BorderLeft =
						boldCellStyle.BorderBottom = BorderStyle.Medium;
			foreach (var cell in ro.Cells)
				cell.CellStyle = boldCellStyle;
			IFont fontRegular = hssfworkbook.CreateFont();
			fontRegular.Boldweight = (short)FontBoldWeight.Normal;
			fontRegular.FontHeightInPoints = 10;
			fontRegular.FontName = "Arial";
			fontRegular.IsItalic = false;
			var regularCellStyle = hssfworkbook.CreateCellStyle();
			regularCellStyle.SetFont(fontRegular);
			regularCellStyle.WrapText = true;
			regularCellStyle.BorderRight =
				regularCellStyle.BorderTop =
					regularCellStyle.BorderLeft =
						regularCellStyle.BorderBottom = BorderStyle.Thin;
			int totalRequests = 0, solvedRequests = 0;
			var normativesList = NormativeController.List();
			var periodList = PeriodController.List();
			foreach (object o in requests)
			{
				var request = o as REQUEST;
				if (request == null)
					continue;
				i++;
				solvedRequests += (request.SOLVED.HasValue && request.SOLVED.Value) ? 1 : 0;
				var ro1 = sheet1.CreateRow(i);
				ro1.CreateCell(0).SetCellValue(request.USER.NAME);
				ro1.CreateCell(1).SetCellValue(request.POST.NAME);
				ro1.CreateCell(2).SetCellValue(request.DATECREATED.AddHours(request.USER.POST.regions.timezone).ToString());


				//опс
				//entities.Entry(request.DEVICEINFOS.ElementAt(k)).Reference(x => x.OPS).Load();
				int deviceCount = request.DEVICEINFOS.Count;
				if (deviceCount > 1)
				{
					ro1.CreateCell(3).SetCellValue("Групповая заявка");
					ro1.CreateCell(4).SetCellValue("-");
					ro1.CreateCell(5).SetCellValue("-");
					ro1.CreateCell(6).SetCellValue("-");
				}
				else
				{
					ro1.CreateCell(3).SetCellValue(string.Format("{0}, {1}", request.DEVICEINFOS.First().OPS.idx,
						request.DEVICEINFOS.First().OPS.name_ops));

					ro1.CreateCell(4).SetCellValue(string.Format("{0}", request.DEVICEINFOS.First().OPS.@class));

					ro1.CreateCell(5).SetCellValue(string.Format("{0}, {1}", request.DEVICEINFOS.First().type,
						request.DEVICEINFOS.First().DEVICEMODEL));
					ro1.CreateCell(6).SetCellValue(request.DEVICEINFOS.First().DEVICENUMBER);
				}
				ro1.CreateCell(7).SetCellValue(request.services != null ? request.services.name : "");
				ro1.CreateCell(8).SetCellValue(request.TEXT);
				ro1.CreateCell(9).SetCellValue(request.SOLVED.HasValue && request.SOLVED.Value ? "Да" : "Нет");

				ro1.CreateCell(10).SetCellValue(request.BF_NUMBER);
				ro1.CreateCell(11).SetCellValue(request.BF_STATUS == null ? string.Empty : request.BF_STATUS.TEXT);
				ro1.CreateCell(12).SetCellValue(request.BF_COMMENT);


				ro1.CreateCell(14).SetCellValue(request.NUMBER);
				StringBuilder solvePeriod = new StringBuilder();

				if (request.SOLVED.HasValue && request.SOLVED.Value)
				{
					DateTime realCloseDate = request.DATESOLVED.Value.AddHours(request.USER.POST.regions.timezone);
					if (request.act_date.HasValue)
						realCloseDate = request.act_date.Value.AddHours(request.USER.POST.regions.timezone);
					ro1.CreateCell(13).SetCellValue(realCloseDate.ToString());
					TimeSpan ts = request.DATESOLVED.Value.Subtract(request.DATECREATED);
					if (ts.Days > 0)
						solvePeriod.AppendFormat("{0} сут. ", ts.Days);
					if (ts.Hours > 0)
						solvePeriod.AppendFormat("{0} ч. ", ts.Hours);
					if (ts.Minutes > 0)
						solvePeriod.AppendFormat("{0} мин. ", ts.Minutes);
					ro1.CreateCell(15).SetCellValue(solvePeriod.ToString());
				}
				if ((int)parametersValues["postId"] == 656700)
				{
					string message = "";
					string normTime = "";
					var norma = RequestController.SelectNormative(request, out message, normativesList);
					if (norma == null)
						normTime = message;
					else
					{
						normTime = string.Format("{0} ч.", norma.time);
						ro1.CreateCell(16).SetCellValue(normTime);


						DateTime normativeDate = RequestController.SelectNormativeDate(request, norma, out message, periodList);
						string s = "";
						if (!string.IsNullOrEmpty(message))
						{
							s = message;
						}
						else
						{
							if (request.SOLVED.HasValue && request.SOLVED.Value && request.BF_STATUS_ID.HasValue &&
								request.BF_STATUS_ID.Value != BfStatusController.FindByText("Отменен"))
							{
								DateTime realCloseDate = request.DATESOLVED.Value.AddHours(request.USER.POST.regions.timezone);
								if (request.act_date.HasValue)
									realCloseDate = request.act_date.Value.AddHours(request.USER.POST.regions.timezone);
								if (realCloseDate > normativeDate)
									s = "Просрочена";
								else
									s = "Не просрочена";
							}
							else
							{
								//TODO Учесть часовой пояс
								if (DateTime.Now.AddHours(request.USER.POST.regions.timezone) >= normativeDate)
								{
									s = "Просрочена";
								}
								else
								{
									s = "Не просрочена";
								}
							}
						}
						ro1.CreateCell(17).SetCellValue(s);
					}
				}
				foreach (ICell cell in ro1.Cells)
					cell.CellStyle = regularCellStyle;
			}
			totalRequests = i;
			var cell1 = sheet1.CreateRow(i + 1).CreateCell(11);
			cell1.SetCellValue(totalRequests.ToString());
			cell1.CellStyle = boldCellStyle;

			cell1 = sheet1.GetRow(i + 1).CreateCell(7);
			cell1.SetCellValue("Закрыто:");
			cell1.CellStyle = boldCellStyle;

			cell1 = sheet1.GetRow(i + 1).CreateCell(10);
			cell1.SetCellValue("Всего:");
			cell1.CellStyle = boldCellStyle;


			cell1 = sheet1.GetRow(i + 1).CreateCell(8);
			cell1.SetCellValue(solvedRequests.ToString());
			cell1.CellStyle = boldCellStyle;

			cell1 = sheet1.CreateRow(i + 2).CreateCell(8);
			cell1.SetCellValue("Автор");
			cell1.CellStyle = boldCellStyle;

			cell1 = sheet1.GetRow(i + 2).CreateCell(9);
			cell1.CellStyle = boldCellStyle;
			int authorId = (int)parametersValues["userId"];
			if (authorId == 0)
				cell1.SetCellValue("Все");
			else
				cell1.SetCellValue(UserController.GetUserById(authorId).NAME);

			cell1 = sheet1.CreateRow(i + 3).CreateCell(8);
			cell1.SetCellValue("Почтамт");
			cell1.CellStyle = boldCellStyle;

			cell1 = sheet1.GetRow(i + 3).CreateCell(9);
			cell1.CellStyle = boldCellStyle;
			int postId = (int)parametersValues["postIdFilter"];
			if (postId == 0)
				cell1.SetCellValue("Все");
			else
				cell1.SetCellValue(PostController.FindById(postId).NAME);

			cell1 = sheet1.CreateRow(i + 4).CreateCell(8);
			cell1.SetCellValue("Фильтр по индексу");
			cell1.CellStyle = boldCellStyle;

			cell1 = sheet1.GetRow(i + 4).CreateCell(9);
			cell1.CellStyle = boldCellStyle;
			var opsId = parametersValues["opsId"];
			if (opsId == null)
				cell1.SetCellValue("Все");
			else
				cell1.SetCellValue(opsId.ToString());

			cell1 = sheet1.CreateRow(i + 5).CreateCell(8);
			cell1.SetCellValue("Фильтр по модели");
			cell1.CellStyle = boldCellStyle;

			cell1 = sheet1.GetRow(i + 5).CreateCell(9);
			cell1.CellStyle = boldCellStyle;
			var model = parametersValues["deviceModel"];
			if (model == null)
				cell1.SetCellValue("Все");
			else
				cell1.SetCellValue(model.ToString());

			sheet1.CreateFreezePane(0, 1);
			List<double> columnWidths = new List<double>();
			columnWidths.Add(14);    // 1
			columnWidths.Add(47.86); // 2
			columnWidths.Add(17.71); // 3
			columnWidths.Add(35.86); // 4
			columnWidths.Add(6.29);  // 5
			columnWidths.Add(43.29); // 6
			columnWidths.Add(34.43); // 7
			columnWidths.Add(15.14); // 8
			columnWidths.Add(35.29); // 9
			columnWidths.Add(15.14); // 10
			columnWidths.Add(17.43); // 11
			columnWidths.Add(14.71); // 12
			columnWidths.Add(44.29); // 13
			columnWidths.Add(17.71); // 14
			columnWidths.Add(11.14); // 15
			columnWidths.Add(18.43); // 16
			columnWidths.Add(11.14); // 17
			columnWidths.Add(11); // 18

			for (int j = 0; j < columnWidths.Count; j++)
			{
				columnWidths[j] = (columnWidths[j] - 1) * 256 + 438;
				sheet1.SetColumnWidth(j, (int)Math.Round(columnWidths[j]));

			}
			/*foreach (var columnCell in sheet1.GetRow(0).Cells)
		{
			sheet1.AutoSizeColumn(columnCell.ColumnIndex);
		}*/
			context.Response.ContentType = "application/octet-stream";
			context.Response.AddHeader("Content-Disposition", "attachment;filename=report.xls");
			context.Response.Clear();
			MemoryStream ms = new MemoryStream();
			hssfworkbook.Write(ms);

			ms.Flush();
			context.Response.BinaryWrite(ms.GetBuffer());
			ms.Dispose();
			context.Response.End();
		}

		public static IList<REQUEST> FindByOps(int id)
		{
			using (hdEntities entities = new hdEntities())
			{
				var r = entities.REQUEST.Include("POST").Include("USER").Include("BF_STATUS").Include("DEVICEINFOS").ToList();
				return r.Where(x => x.DEVICEINFOS.ElementAt(0).OPS_ID == id).ToList();
			}
		}

		public static void LoadFiles(REQUEST request)
		{
			using (hdEntities entities = new hdEntities())
			{
				entities.Attach(request);
				//entities.Entry(request).Collection(x => x.files).Load();
				entities.LoadProperty(request, x => x.files);
			}
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
			public normatives norma;
			public DateTime normativeDate;

			public RequestData(DateTime? closeDate, DateTime? actCloseDate, string penaltyType, string comment, normatives norma, DateTime normativeDate)
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

			using (FileStream fs = new FileStream(Path.Combine(templateDirectory, "penalties.xls"), FileMode.Open))
			{
				HSSFWorkbook workbook = new HSSFWorkbook(fs);
				HSSFSheet activeSheet = (HSSFSheet)workbook.GetSheetAt(0);
				int startRow = activeSheet.FirstRowNum + 10;
				POST post = PostController.FindById(postId);
				var r = new RequestController().FindByPost(post.ID);
				var opses = new OpsController().List(post.ID, OpsController.OpsFilter.All);
				int i = 0, rowIndex = startRow;
				DateTime periodStart = new DateTime(currentPeriodYear, currentPeriodMonth, 1);
				DateTime periodFinish = new DateTime(currentPeriodYear, currentPeriodMonth,
					DateTime.DaysInMonth(currentPeriodYear, currentPeriodMonth)).AddDays(1).AddSeconds(-1);
				var row1 = activeSheet.GetRow(2);
				row1.Cells[1].SetCellValue(string.Format("{0:d} - {1:d}", periodStart, periodFinish));

				if (!post.regionsReference.IsLoaded)
					PostController.LoadRegion(post);
				if (post.regions.region_coeff.HasValue && post.regions.equip_coeff.HasValue)
				{
					activeSheet.GetRow(0).Cells[1].SetCellValue(post.regions.region_coeff.Value);
					activeSheet.GetRow(1).Cells[1].SetCellValue(post.regions.equip_coeff.Value);
				}
				int softwareServiceId = ServiceController.SelectIdByText("Обновление, установка и настройка ПО");
				List<KeyValuePair<REQUEST, RequestData>> requestDatas = new List<KeyValuePair<REQUEST, RequestData>>();
				List<normatives> normativesList = NormativeController.List();
				IList<periods> periodList = PeriodController.List();
				foreach (OPS ops in opses)
				{
					//entities.Attach(south);
					var reqs = r.Where(x => x.DEVICEINFOS.First().OPS_ID == ops.id).ToList();
					int заявштр1 = 0;
					int заявштр2 = 0;
					int заявштр3 = 0;
					int заявштр4 = 0;
					int заявсвоевр1 = 0;
					int заявсвоевр2 = 0;

					double качествоSLA = 85;

					List<REQUEST> заявштр1List = new List<REQUEST>();
					List<REQUEST> заявштр2List = new List<REQUEST>();
					List<REQUEST> заявштр3List = new List<REQUEST>();
					List<REQUEST> заявштр4List = new List<REQUEST>();
					List<REQUEST> заявсвоевр1List = new List<REQUEST>();
					List<REQUEST> заявсвоевр2List = new List<REQUEST>();

					if (!considerActs)
						getOpsPenalties(reqs, periodStart, periodFinish, softwareServiceId, ref заявштр1, ref заявштр2, ref заявштр3, ref заявштр4, ref заявсвоевр1, ref заявсвоевр2,
							заявштр1List, заявштр2List, заявштр3List, заявштр4List, заявсвоевр1List, заявсвоевр2List, requestDatas, normativesList, periodList);
					else
						getOpsPenaltiesConsideringActs(reqs, periodStart, periodFinish, softwareServiceId, ref заявштр1, ref заявштр2, ref заявштр3, ref заявштр4, ref заявсвоевр1, ref заявсвоевр2,
							заявштр1List, заявштр2List, заявштр3List, заявштр4List, заявсвоевр1List, заявсвоевр2List, requestDatas, normativesList, periodList);

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
						row.Cells[startColumnIndex].SetCellValue(ops.idx + " " + ops.name_ops);
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
						switch (ops.@class)
						{
							case 101:
								СТаб = 203300;
								classString = "АУ филиала I категории";
								break;
							case 102:
								СТаб = 66300;
								classString = "АУ филиала II категории";
								break;
							case 103:
								СТаб = 49100;
								classString = "АУ филиала III категории";
								break;
							case 104:
								СТаб = 35000;
								classString = "АУ филиала IV категории";
								break;
							case 11:
								СТаб = 58650;
								classString = "АУ почтамта I категории";
								break;
							case 12:
								СТаб = 30100;
								classString = "АУ почтамта II категории";
								break;
							case 13:
								СТаб = 17900;
								classString = "АУ почтамта III категории";
								break;
							case 14:
								СТаб = 11700;
								classString = "АУ почтамта IV категории";
								break;
							case 15:
								СТаб = 34910;
								classString = "УКД";
								break;
							case 1:
								СТаб = 66950;
								classString = "ОПС класс 1";
								break;
							case 2:
								СТаб = 44990;
								classString = "ОПС класс 2";
								break;
							case 3:
								СТаб = 13640;
								classString = "ОПС класс 3";
								break;
							case 4:
								СТаб = 5130;
								classString = "ОПС класс 4";
								break;
							case 5:
								СТаб = 3190;
								classString = "ОПС класс 5";
								break;
							default:
								classString = "Не  указан класс ОПС";
								break;
						}
						if (ops.name_ops.ToLower().Contains("попс") || ops.type_city == "2")
						{
							СТаб = 3780;
							classString = "ПОПС";
						}
						row.Cells[startColumnIndex + 8].SetCellValue(кштр);
						row.Cells[startColumnIndex + 9].SetCellValue(classString);
						row.Cells[startColumnIndex + 10].SetCellValue(СТаб);
						string abonCell = CellReference.ConvertNumToColString(startColumnIndex + 11 - 1) + (row.RowNum + 1);
						string penaltyCoefCell = CellReference.ConvertNumToColString(startColumnIndex + 11 - 3) + (row.RowNum + 1);
						string qualityCell = CellReference.ConvertNumToColString(startColumnIndex + 11 - 4) + (row.RowNum + 1);
						row.Cells[startColumnIndex + 11].SetCellFormula(abonCell + "*max(85-" + qualityCell + ",0)/100*" + penaltyCoefCell + "*$B$1*$B$2");
						//row.Cells[startColumnIndex+11].r;
						rowIndex++;
					}
				}
				requestDatas.Sort(
					(x, y) => x.Key.DATECREATED.CompareTo(y.Key.DATECREATED));
				ISheet secondSheet = workbook.GetSheetAt(1);
				int firstrow = secondSheet.FirstRowNum + 1;
				for (int index = 0; index < requestDatas.Count; index++)
				{
					var requestData = requestDatas[index];
					IRow currentRow = secondSheet.CreateRow(firstrow + index);
					int firstcolumn = 0;
					//номер заявки
					currentRow.CreateCell(firstcolumn).SetCellValue(requestData.Key.BF_NUMBER);
					//ОПС
					currentRow.CreateCell(firstcolumn + 1).SetCellValue(requestData.Key.DEVICEINFOS.First().OPS.idx + "_" + requestData.Key.DEVICEINFOS.First().OPS.name_ops);
					//вид услуги
					currentRow.CreateCell(firstcolumn + 2).SetCellValue(requestData.Key.services.name);
					// статус
					currentRow.CreateCell(firstcolumn + 3).SetCellValue(requestData.Key.BF_STATUS.TEXT);
					// дата заявки
					currentRow.CreateCell(firstcolumn + 4).SetCellValue(requestData.Key.DATECREATED.AddHours(requestData.Key.USER.POST.regions.timezone).ToString());
					var ts = DateTime.Now - requestData.Key.DATECREATED;
					if (requestData.Value.actCloseDate.HasValue)
						ts = requestData.Value.actCloseDate.Value -
							 requestData.Key.DATECREATED.AddHours(requestData.Key.USER.POST.regions.timezone);
					else
						if (requestData.Value.closeDate.HasValue)
						ts = requestData.Value.closeDate.Value -
						 requestData.Key.DATECREATED.AddHours(requestData.Key.USER.POST.regions.timezone);
					// дни
					currentRow.CreateCell(firstcolumn + 5).SetCellValue(ts.Days);
					// часы
					currentRow.CreateCell(firstcolumn + 6).SetCellValue(ts.Subtract(new TimeSpan(ts.Days, 0, 0, 0)).ToString(@"hh\:mm\:ss"));
					// дата закрытия
					currentRow.CreateCell(firstcolumn + 7).SetCellValue(requestData.Value.closeDate.HasValue ? requestData.Value.closeDate.Value.ToString() : string.Empty);
					// дата закрытия + наличие акта
					currentRow.CreateCell(firstcolumn + 8).SetCellValue(requestData.Value.actCloseDate.HasValue ? requestData.Value.actCloseDate.Value.ToString() : string.Empty);
					// срок ремонта
					if (requestData.Value.norma.dayFlag.HasValue && requestData.Value.norma.dayFlag.Value)
						currentRow.CreateCell(firstcolumn + 9).SetCellValue(requestData.Value.norma.time + " д.");
					else
						currentRow.CreateCell(firstcolumn + 9).SetCellValue(requestData.Value.norma.time + " ч.");
					// комментарий
					currentRow.CreateCell(firstcolumn + 10).SetCellValue(requestData.Value.comment);
					// планируемая дата закрытия
					currentRow.CreateCell(firstcolumn + 11).SetCellValue(requestData.Value.normativeDate.ToString());
					//штрафы
					currentRow.CreateCell(firstcolumn + 12).SetCellValue(requestData.Value.penaltyType);
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

		public static void getOpsPenalties(List<REQUEST> reqs, DateTime periodStart, DateTime periodFinish, int softwareServiceId, ref int заявштр1, ref int заявштр2, ref int заявштр3, ref int заявштр4, ref int заявсвоевр1, ref int заявсвоевр2, List<REQUEST> заявштр1List, List<REQUEST> заявштр2List, List<REQUEST> заявштр3List, List<REQUEST> заявштр4List, List<REQUEST> заявсвоевр1List, List<REQUEST> заявсвоевр2List, List<KeyValuePair<REQUEST, RequestData>> requestDatas, List<normatives> normativesList, IList<periods> periodList)
		{
			foreach (var request in reqs)
			{
				DateTime requestDate = request.DATECREATED.AddHours(request.USER.POST.regions.timezone);
				string message;
				var norma = RequestController.SelectNormative(request, out message, normativesList);
				if (norma == null)
					continue;
				// целевая дата в местном часовом поясе
				DateTime normativeDate = RequestController.SelectNormativeDate(request, norma, out message, periodList);

				if (request.BF_STATUS == null)
					continue;
				if (request.BF_STATUS.TEXT == "Отменен")
					continue;
				if (request.BF_STATUS.TEXT == "Назначен")
				{
					// неисполненные с целевой датой в прошлом периоде
					if (normativeDate < periodStart)
					{
						// не включать если настройка и обновление ПО
						if (request.service_id != softwareServiceId)
						{
							заявштр1++;
							заявштр1List.Add(request);
							requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
								new RequestData(null, null, "ЗАЯВштр1", "Не выполнена." + message, norma, normativeDate)));
						}
					}
					else
					// неисполненные с целевой датой в текущем периоде
					if (periodStart <= normativeDate && normativeDate <= periodFinish)
					{
						заявштр2++;
						заявштр2List.Add(request);
						requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
							new RequestData(null, null, "ЗАЯВштр2", "Не выполнена." + message, norma, normativeDate)));
					}
				}
				else if (request.BF_STATUS.TEXT == "Закрыт")
				{
					DateTime realCloseDate = request.DATESOLVED.Value.AddHours(request.USER.POST.regions.timezone);
					DateTime closeDate = realCloseDate;
					DateTime? actCloseDate = null;
					if (request.act_date.HasValue)
					{
						realCloseDate = request.act_date.Value.AddHours(request.USER.POST.regions.timezone);
						actCloseDate = realCloseDate;
					}
					// закрытые в текущем периоде с целевой датой в прошлом периоде
					if (periodStart <= realCloseDate && realCloseDate <= periodFinish && normativeDate < periodStart)
					{
						// не включать если настройка и обновление ПО
						if (request.service_id != softwareServiceId)
						{
							заявштр3++;
							заявштр3List.Add(request);
							requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
								new RequestData(closeDate, actCloseDate, "ЗАЯВштр3", "Выполнена не в срок." + message, norma, normativeDate)));
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
						requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
							new RequestData(closeDate, actCloseDate, "ЗАЯВштр4", "Выполнена не в срок." + message, norma, normativeDate)));
					}
					else
					// закрытые в текущем периоде с целевой датой в текущем периоде с выполнением в срок
					if (periodStart <= realCloseDate && realCloseDate <= periodFinish && periodStart <= normativeDate &&
							 normativeDate <= periodFinish &&
							 realCloseDate <= normativeDate)
					{
						заявсвоевр1++;
						заявсвоевр1List.Add(request);
						requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
							new RequestData(closeDate, actCloseDate, "ЗАЯВсвоевр1", "Выполнена в срок." + message, norma, normativeDate)));
					}
					else
					// закрытые в текущем периоде с целевой датой в следующем периоде
					if (periodStart <= realCloseDate && realCloseDate <= periodFinish && normativeDate > periodFinish &&
							 realCloseDate <= normativeDate)
					{
						заявсвоевр2++;
						заявсвоевр2List.Add(request);
						requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
							new RequestData(closeDate, actCloseDate, "ЗАЯВсвоевр2", "Выполнена в срок." + message, norma, normativeDate)));
					}
				}
			}
		}

		public static void getOpsPenaltiesConsideringActs(List<REQUEST> reqs, DateTime periodStart, DateTime periodFinish, int softwareServiceId, ref int заявштр1, ref int заявштр2, ref int заявштр3, ref int заявштр4, ref int заявсвоевр1, ref int заявсвоевр2, List<REQUEST> заявштр1List, List<REQUEST> заявштр2List, List<REQUEST> заявштр3List, List<REQUEST> заявштр4List, List<REQUEST> заявсвоевр1List, List<REQUEST> заявсвоевр2List, List<KeyValuePair<REQUEST, RequestData>> requestDatas, List<normatives> normativesList, IList<periods> periodList)
		{
			foreach (var request in reqs)
			{
				DateTime requestDate = request.DATECREATED.AddHours(request.USER.POST.regions.timezone);
				string message;
				var norma = RequestController.SelectNormative(request, out message, normativesList);
				if (norma == null)
					continue;
				// целевая дата в местном часовом поясе
				DateTime normativeDate = RequestController.SelectNormativeDate(request, norma, out message, periodList);
				if (request.BF_STATUS == null)
					continue;
				if (request.BF_STATUS.TEXT == "Отменен")
					continue;
				if (request.BF_STATUS.TEXT == "Назначен")
				{
					// неисполненные с целевой датой в прошлом периоде
					if (normativeDate < periodStart)
					{
						// не включать если настройка и обновление ПО
						if (request.service_id != softwareServiceId)
						{
							заявштр1++;
							заявштр1List.Add(request);
							requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
								new RequestData(null, null, "ЗАЯВштр1", "Не выполнена." + message, norma, normativeDate)));
						}
					}
					else
					// неисполненные с целевой датой в текущем периоде
					if (periodStart <= normativeDate && normativeDate <= periodFinish)
					{
						заявштр2++;
						заявштр2List.Add(request);
						requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
							new RequestData(null, null, "ЗАЯВштр2", "Не выполнена." + message, norma, normativeDate)));
					}
				}
				else if (request.BF_STATUS.TEXT == "Закрыт")
				{
					DateTime realCloseDate = request.DATESOLVED.Value.AddHours(request.USER.POST.regions.timezone);
					DateTime closeDate = realCloseDate;
					// если заявка закрыта, но нет акта, считать неисполненной
					if (!request.act_date.HasValue)
					{
						// неисполненные с целевой датой в прошлом периоде
						if (normativeDate < periodStart)
						{
							// не включать если настройка и обновление ПО
							if (request.service_id != softwareServiceId)
							{
								заявштр1++;
								заявштр1List.Add(request);
								requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
									new RequestData(closeDate, null, "ЗАЯВштр1", "Не выполнена." + message, norma, normativeDate)));
							}
						}
						else
						// неисполненные с целевой датой в текущем периоде
						if (periodStart <= normativeDate && normativeDate <= periodFinish)
						{
							заявштр2++;
							заявштр2List.Add(request);
							requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
								new RequestData(closeDate, null, "ЗАЯВштр2", "Не выполнена." + message, norma, normativeDate)));
						}
					}
					else
					{
						realCloseDate = request.act_date.Value.AddHours(request.USER.POST.regions.timezone);
						DateTime actCloseDate = realCloseDate;
						// закрытые в текущем периоде с целевой датой в прошлом периоде
						if (periodStart <= realCloseDate && realCloseDate <= periodFinish && normativeDate < periodStart)
						{
							// не включать если настройка и обновление ПО
							if (request.service_id != softwareServiceId)
							{
								заявштр3++;
								заявштр3List.Add(request);
								requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
									new RequestData(closeDate, actCloseDate, "ЗАЯВштр3", "Выполнена не в срок." + message, norma, normativeDate)));
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
							requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
								new RequestData(closeDate, actCloseDate, "ЗАЯВштр4", "Выполнена не в срок." + message, norma, normativeDate)));
						}
						else
						// закрытые в текущем периоде с целевой датой в текущем периоде с выполнением в срок
						if (periodStart <= realCloseDate && realCloseDate <= periodFinish && periodStart <= normativeDate &&
								 normativeDate <= periodFinish &&
								 realCloseDate <= normativeDate)
						{
							заявсвоевр1++;
							заявсвоевр1List.Add(request);
							requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
								new RequestData(closeDate, actCloseDate, "ЗАЯВсвоевр1", "Выполнена в срок." + message, norma, normativeDate)));
						}
						else
						// закрытые в текущем периоде с целевой датой в следующем периоде
						if (periodStart <= realCloseDate && realCloseDate <= periodFinish && normativeDate > periodFinish &&
								 realCloseDate <= normativeDate)
						{
							заявсвоевр2++;
							заявсвоевр2List.Add(request);
							requestDatas.Add(new KeyValuePair<REQUEST, RequestData>(request,
								new RequestData(closeDate, actCloseDate, "ЗАЯВсвоевр2", "Выполнена в срок." + message, norma, normativeDate)));
						}
					}
				}
			}
		}

		public static string SaveReportToDisk(byte[] buffer, string storageRoot, int id, string baseName)
		{
			if (!Directory.Exists(Path.Combine(storageRoot, id.ToString())))
				Directory.CreateDirectory(Path.Combine(storageRoot, id.ToString()));
			string userStorageRoot = Path.Combine(storageRoot, id.ToString());
			string reportFileName = string.Format("{1}_{0}.xls", Guid.NewGuid(), baseName);
			File.WriteAllBytes(Path.Combine(userStorageRoot, reportFileName), buffer);
			return Path.Combine(id.ToString(), reportFileName);
		}

		public static void InsertFile(REQUEST request, files file)
		{
			using (hdEntities entities = new hdEntities())
			{
				entities.Attach(file);
				entities.Attach(request);
				request.files.Add(file);
				entities.SaveChanges();
			}
		}
		static string ArrReplace(string source, char[] invalidChars)
		{
			System.Array.ForEach(invalidChars, invalidChar => source = source.Replace(invalidChar, '_'));
			return source;
		}

		public static void SaveEmailAsFileToRequestAndDisk(REQUEST request, DateTime dateTimeCreated, Item message, string mailSubjectString, string messageId)
		{
			RequestController.LoadFiles(request);
			string fileName = "storage\\emails\\" + dateTimeCreated.ToString("yyyy_MM_dd_HH_mm_ss_")
					+ ArrReplace(mailSubjectString, Path.GetInvalidFileNameChars()).Substring(0, Math.Min(100, mailSubjectString.Length))
					+ ".eml";
			if (request.files.All(x => x.message_id != messageId && x.fileUrl != fileName))
			{
				string storageRootUrl = ParameterController.FindByName("storageRootUrl").value;

				string filePath = Path.Combine(storageRootUrl, fileName);
				using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
				{
					message.Load(new Microsoft.Exchange.WebServices.Data.PropertySet(ItemSchema.MimeContent));
					fileStream.Write(message.MimeContent.Content, 0, message.MimeContent.Content.Length);
				}
				files file = new files()
				{
					message_id = messageId,
					contentType = "eml",
					fileUrl = fileName
				};
				FileController.Save(file);
				RequestController.InsertFile(request, file);
				WebLogger.GetInstance().Info("Добавлен файл {0} для заявки {1} из письма {2}", fileName, request.ID, mailSubjectString);
			}
		}

		public static DEVICEINFO FindChildDeviceInfoBySerialActive(REQUEST parentRequest, List<string> childSerialNumbers, string childOpsActiveName)
		{
			using (hdEntities entities = new hdEntities())
			{
				foreach (var childSerialNumber in childSerialNumbers)
				{
					for (int i = 0; i < parentRequest.DEVICEINFOS.Count; i++)
					{
						DEVICEINFO deviceinfo = parentRequest.DEVICEINFOS.ElementAt(i);
						entities.Attach(deviceinfo);
						entities.LoadProperty(deviceinfo, x => x.OPS);
						if (deviceinfo.DEVICENUMBER == childSerialNumber && deviceinfo.OPS.active_name == childOpsActiveName)
							return deviceinfo;
					}
				}
				return null;
			}
		}

		public static REQUEST CopyToChild(REQUEST parentRequest)
		{
			REQUEST childRequest = new REQUEST();
			childRequest.TEXT = parentRequest.TEXT;
			childRequest.AUTHOR_ID = parentRequest.AUTHOR_ID;
			childRequest.POST_ID = parentRequest.POST_ID;
			childRequest.CONTACT = parentRequest.CONTACT;
			childRequest.SHORTTEXT = parentRequest.SHORTTEXT;
			childRequest.service_id = parentRequest.service_id;
			childRequest.parent_id = parentRequest.ID;
			childRequest.DATECREATED = parentRequest.DATECREATED;
			foreach (files file in parentRequest.files)
			{
				files file1 = new files()
				{
					contentType = file.contentType,
					fileUrl = file.fileUrl,
					message_id = file.message_id
				};

				childRequest.files.Add(file1);
			}
			return childRequest;
		}

		public static DEVICEINFO FindChildDeviceInfoBySerialIndex(REQUEST parentRequest, List<string> childSerialNumbers, string idx)
		{
			using (hdEntities entities = new hdEntities())
			{
				foreach (var childSerialNumber in childSerialNumbers)
				{
					for (int i = 0; i < parentRequest.DEVICEINFOS.Count; i++)
					{
						DEVICEINFO deviceinfo = parentRequest.DEVICEINFOS.ElementAt(i);
						entities.Attach(deviceinfo);
						entities.LoadProperty(deviceinfo, x => x.OPS);
						int iidx;
						if (int.TryParse(idx, out iidx))
						{
							if (deviceinfo.DEVICENUMBER == childSerialNumber && deviceinfo.OPS.idx == iidx)
								return deviceinfo;
						}
					}
				}
				return null;
			}
		}

		public static DEVICEINFO FindChildDeviceInfoByIndex(REQUEST parentRequest, string idx)
		{
			using (hdEntities entities = new hdEntities())
			{
				for (int i = 0; i < parentRequest.DEVICEINFOS.Count; i++)
				{
					DEVICEINFO deviceinfo = parentRequest.DEVICEINFOS.ElementAt(i);
					entities.Attach(deviceinfo);
					entities.LoadProperty(deviceinfo, x => x.OPS);
					int iidx;
					if (int.TryParse(idx, out iidx))
					{
						if (deviceinfo.OPS.idx == iidx)
							return deviceinfo;
					}
				}
				return null;
			}
		}

		public static string GetParentNumber(int id)
		{
			using (hdEntities entities = new hdEntities())
			{
				REQUEST r = entities.REQUEST.SingleOrDefault(x => x.ID == id);
				if (r != null)
					return r.BF_NUMBER;
				return string.Empty;
			}
		}
	}
}
