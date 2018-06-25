using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.ServiceProcess;
using System.Text;
using MYERP_ServerServiceRuner.Base;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;


namespace MYERP_ServerServiceRuner
{
	public partial class MainService
	{
		#region 發送3天排期異常
		void DeliverPlanFinishStatisticErrorSender()
		{
			MyRecord.Say("开启计算3日出货计划异常..........");
			Thread t = new Thread(DeliverPlanFinishStatisticErrorSendMail);
			t.IsBackground = true;
			t.Start();
			MyRecord.Say("定时发送3日出货计划异常已经启动。");
		}

		public static string LCStr(string txt)
		{
			return MyConvert.ZHLC(txt);
		}

		void DeliverPlanFinishStatisticErrorSendMail()
		{
			try
			{
				MyRecord.Say("------------------开始定时发送3日异常报表----------------------------");
				MyRecord.Say("从数据库搜寻内容");
				DateTime NowTime = DateTime.Now;
				string body = @"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#e6f3af #ffffff>
<DIV><FONT size=3 face=PMingLiU>{3}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;</FONT></DIV>
<DIV><FONT color=#ff0000 size=5 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <B>{0:yy/MM/dd HH:mm} {1}</B></FONT></DIV>
{4}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;每日中午11点55发送，提醒各位生管检查排程正确性。</DIV>
<DIV><FONT color=#0000ff size=2 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，切勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{2:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
";
				string gridTitle = @"
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工序
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	料号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	问题描述
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	交期备注
	</TD>
	</TR>
	{0}
</TBODY></TABLE></FONT>
</DIV>";
				string br = @"
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{0}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{1}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{2}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{3}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{4}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=left >
	{5}
	</TD>
	</TR>";

				string mailBodyMainSentence = string.Empty;
				string brs = string.Empty;
				DateTime DBegin = DateTime.Now.AddDays(-1).Date.AddHours(23), DEnd = DateTime.Now.Date.AddDays(3).Date.AddHours(8);
				MyRecord.Say(string.Format("开始检查工单交期在{0:yy/MM/dd HH:mm}-{1:yy/MM/dd HH:mm}之间所有排程有问题的工单。", DBegin, DEnd));
				LoadDataSource(DBegin, DEnd);
				LoadGridSchame(DBegin, DEnd);
				MyRecord.Say("读取计算完成。");

				var vListGridRow = from a in GridList
								   where a.isError
								   select a;
				List<GridProcessItem> AllProcessItemList = new List<GridProcessItem>();
				foreach (var item in vListGridRow)
				{
					AllProcessItemList.AddRange(item.GridProcess);
				}

				var viss = from a in AllProcessItemList
						   where a.isError
						   orderby a.DepartmentFullSortID, a.ProcessCode
						   select a;

				int xProduceGridCount = viss.Count();
				MyRecord.Say("创建SendMail。");
				MyBase.SendMail sm = new MyBase.SendMail();
				string fname = string.Empty;
				if (xProduceGridCount > 0)
				{
					MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", xProduceGridCount));
					mailBodyMainSentence = string.Format("发现工单交期在{1:yy/MM/dd HH:mm}-{2:yy/MM/dd HH:mm}，有{0}条工单排程有问题。详情请查看附档。", xProduceGridCount, DBegin, DEnd);
					string tmpFileName = Path.GetTempFileName();
					MyRecord.Say(string.Format("tmpFileName = {0}", tmpFileName));
					Regex rgx = new System.Text.RegularExpressions.Regex(@"(?<=tmp)(.+)(?=\.tmp)");
					string tmpFileNameLast = rgx.Match(tmpFileName).Value;
					MyRecord.Say(string.Format("tmpFileNameLast = {0}", tmpFileNameLast));
					string dName = string.Format("{0}\\TMPExcel", Application.StartupPath);
					if (!Directory.Exists(dName))
					{
						Directory.CreateDirectory(dName);
					}
					fname = string.Format("{0}\\{1}", dName, string.Format("NOTICE_{0:yyyyMMdd}_tmp{1}.xls", NowTime.Date, tmpFileNameLast));
					MyRecord.Say(string.Format("fname = {0}", fname));
					ExportToExcelDeliveryPlan(fname);
					MyRecord.Say("已经保存了。");
					sm.Attachments.Add(new System.Net.Mail.Attachment(fname));
					MyRecord.Say("加载到附件");
					foreach (var item in viss)
					{
						brs += string.Format(br, item.DepartmentName, item.ProcessName, item.ProductName, item.ProduceRdsNo, item.IssuesMemo, item.SRemark);
					}
					brs = string.Format(gridTitle, brs);
					MyRecord.Say("生成邮件内容。");
				}
				else
				{
					mailBodyMainSentence = "，没有发现3日出货计划和排程异常情况。";
					brs = string.Empty;
				}

				MyRecord.Say("加载邮件内容。");

				sm.MailBodyText = string.Format(body, NowTime, mailBodyMainSentence, DateTime.Now, MyBase.CompanyTitle, brs);
				sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}三日出货计划和排程异常表", NowTime.AddDays(-1).Date, MyBase.CompanyTitle));
				//string mailto = ConfigurationManager.AppSettings["DeliveryPlanMailTo"], 
				//       mailcc = ConfigurationManager.AppSettings["DeliveryPlanMailCC"];
				MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("DeliveryPlan");
				MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
				sm.MailTo = mAddress.MailTo;
				sm.MailCC = mAddress.MailCC;
				//sm.MailTo = "my80@my.imedia.com.tw";
				MyRecord.Say("发送邮件。");
				sm.SendOut();
				sm.mail.Dispose();
				sm = null;
				MyRecord.Say("已经发送。");
				if (File.Exists(fname))
				{
					File.Delete(fname);
				}
				MyRecord.Say("------------------三日排期异常---发送完成----------------------------");
			}
			catch (Exception ex)
			{
				MyRecord.Say(ex);
			}
		}


		protected List<Plan_ProduceNoteItem> MainData, FinishData, PlanData;

		List<GridProduceNoteItem> GridList;

		protected string[] titlewords = new string[] { "序号", "工单号", "产品号", "料号", "订单数", "交期", "需求数", "台号", "内容", "库存数" };
		protected string[] titleKeys = new string[] { "ID", "ProductRdsNo", "ProductCode", "ProductName", "ProductNumb", "SendDate", "pNumb", "PartID", "Title", "StockNumb" };
		/// <summary>
		/// 加载数据源
		/// </summary>
		/// <param name="wp"></param>
		protected virtual void LoadDataSource(DateTime DateBegin, DateTime DateEnd)
		{
			///工单内容
			string SQL = @"
SET NOCOUNT ON

DECLARE @T Table ([ProdRdsNo] NVarchar(40),
				  [ProductID] Int,
				  [SENDDATE] DateTime,
				  [PNUMB] Numeric(38, 2),
				  [Inputer] NVarchar(40),
				  [Inputdate] DateTime,
				  [CustCode] Varchar(16),
				  [ProdCode] Varchar(50),
				  [ProdName] NVarchar(800),
				  [PartID] NVarchar(400),
				  [ProcessCode] Varchar(10) NOT NULL,
				  [ProductNumb] Numeric(8, 0),
				  [ProcessName] Varchar(30),
				  [FinishNumb] Decimal(20, 6),
				  [RejectNumb] Decimal(20, 6),
				  [LossedNumb] Decimal(20, 6),
				  [OverDate] DateTime,
				  [SortID] Varchar(40),
				  [SName] NVarchar(400),
				  [FinishRemark] NVarchar(2000),
				  [CloseDate] DateTime,
				  [Closer] NVarchar(20),
				  [MachineCode] Varchar(40),
				  [MachineName] NVarchar(100),
				  [DepartmentName] Varchar(20),
				  [Property] Int,
				  [Status] Int,
				  [StockStats] Int,
				  [SRemark] nvarchar(2000),
				  [DepartmentFullSortID] VarChar(40)
);
INSERT INTO @T ([ProdRdsNo], [ProductID], [SENDDATE], [PNUMB], [Inputer], [Inputdate], [CustCode], [ProdCode], [ProdName], [PartID], [ProcessCode],
				[ProductNumb], [ProcessName], [FinishNumb], [RejectNumb], [LossedNumb], [OverDate], [SortID], [FinishRemark], [CloseDate], [Closer],
				[MachineCode], [MachineName], [DepartmentName], [Property], [Status], [StockStats], [SRemark], [DepartmentFullSortID])
SELECT [a].[PRODNO] AS [ProdRdsNo], [a].[PRODID] AS [ProductID], [a].[SENDDATE], [a].[PNUMB], [a].[Inputer], [a].[Inputdate], [b].[custid] AS [CustCode],
	   [b].[code] AS [ProdCode], [e].[Name] AS [ProdName], (CASE WHEN [c].[PartID] = '' THEN '--' ELSE [c].[PartID] END) AS [PartID],
	   [c].[procno] AS [ProcessCode], [b].[pnumb] AS [ProductNumb], [f].[name] AS [ProcessName], [c].[FinishNumb], [c].[RejectNumb], [c].[LossedNumb],
	   [c].[OverDate], (CASE WHEN ISNULL(c.PartID, '') = '' THEN 110 ELSE 100 END) AS [SortID], [b].[FinishRemark], [c].[CloseDate], [c].[Closer],
	   [m].[code] AS [MachineCode], [m].[name] AS [MachineName], [p].[name] AS [DepartmentName], [b].[property], [b].[status], [b].[StockStatus],[b].[SRemark],
	   [p].[FullSortID]
  FROM [_PMC_DeliverPlan_Sendlist] [a]
	   INNER JOIN [moProduce] [b] ON [a].[PRODNO] = [b].[rdsno]
	   INNER JOIN [moProdProcedure] [c] ON [c].[zbid] = [b].[id]
	   INNER JOIN [AllMaterialView] [e] ON [b].[code] = [e].[code]
	   INNER JOIN [moProcedure] [f] ON [c].[procno] = [f].[code]
	   INNER JOIN [moMachine] [m] ON [c].[machinID] = [m].[code]
	   INNER JOIN [pbDept] [p] ON [m].[DepartmentID] = [p].[_ID]
 WHERE ISNULL([b].[status], 0) > 0
	   AND [b].[CloseDate] IS NULL
	   AND [b].[finishdate] IS NULL
	   AND [b].[StockDate] IS NULL
	   AND ISNULL([b].[finalized], 0) = 0
	   AND ([e].[type] NOT IN ({0}))
	   AND [a].[SENDDATE] BETWEEN @ProdBegin AND @ProdEnd
	   AND [b].[inputdate] < @ProdBegin

INSERT INTO @T ([ProdRdsNo], [ProductID], [SENDDATE], [PNUMB], [Inputer], [Inputdate], [CustCode], [ProdCode], [ProdName], [PartID], [ProductNumb],
				[FinishNumb], [OverDate], [SortID], [FinishRemark], [ProcessCode], [ProcessName], [DepartmentName], [Property], [Status], [StockStats],[SRemark])
SELECT [a].[PRODNO] AS [ProdRdsNo], [a].[PRODID] AS [ProductID], [a].[SENDDATE], [a].[PNUMB], [a].[Inputer], [a].[Inputdate], [b].[custid] AS [CustCode],
	   [b].[code] AS [ProdCode], [e].[Name] AS [ProdName], '--' AS [PartID], [b].[pnumb] AS [ProductNumb], [b].[InStockFinishNumb], [b].[finishdate],
	   120 AS [SortID], [b].[FinishRemark], '9999', '成品入庫', '倉庫', [b].[property], [b].[status], [b].[StockStatus],[b].[SRemark]
  FROM [_PMC_DeliverPlan_Sendlist] [a]
	   INNER JOIN [moProduce] [b] ON [a].[PRODNO] = [b].[rdsno]
	   INNER JOIN [AllMaterialView] [e] ON [b].[code] = [e].[code]
 WHERE ISNULL([b].[StopStock], 0) = 0
	   AND ISNULL([b].[status], 0) > 0
	   AND [b].[CloseDate] IS NULL
	   AND [b].[finishdate] IS NULL
	   AND [b].[StockDate] IS NULL
	   AND ISNULL([b].[finalized], 0) = 0
	   AND ([e].[type] NOT IN ({0}))
	   AND [a].[SENDDATE] BETWEEN @ProdBegin AND @ProdEnd
	   AND [b].[inputdate] < @ProdBegin

UPDATE a
   SET [a].[SName] = ISNULL([P1].[Name], '') + ',' + ISNULL([P2].[Name], '') + ',' + ISNULL([P3].[Name], '')
  FROM @T a
	   LEFT OUTER JOIN
	   (SELECT [sy1].[Name], [sy1].[StatusID]
		  FROM [_SY_Status] [sy1]
		 WHERE [sy1].[Type] = 'POProperty') [P1] ON [P1].[StatusID] = [a].[Property]
	   LEFT OUTER JOIN
	   (SELECT [sy2].[Name], [sy2].[StatusID]
		  FROM [_SY_Status] [sy2]
		 WHERE [sy2].[Type] = 'PO') [P2] ON [P2].[StatusID] = [a].[Status]
	   LEFT OUTER JOIN
	   (SELECT [sy3].[Name], [sy3].[StatusID]
		  FROM [_SY_Status] [sy3]
		 WHERE [sy3].[Type] = 'POStatus') [P3] ON [P3].[StatusID] = [a].[StockStats];

SELECT *
  FROM @T
 ORDER BY [SENDDATE],[ProdRdsNo], [SortID], [PartID];
";
			//SQL = string.Format(SQL, ConfigurationManager.AppSettings["DeliveryPlanFinishStaticExceptProdTypes"]);
			string xFilterString = MyConfig.ApplicationConfig.DeliveryPlanFinishStaticExceptProdTypes;
			SQL = string.Format(SQL, xFilterString);
			MyData.MyDataParameter[] mp = new MyData.MyDataParameter[]
			{
				new MyData.MyDataParameter("@ProdBegin",DateBegin, MyData.MyDataParameter.MyDataType.DateTime),
				new MyData.MyDataParameter("@ProdEnd",DateEnd, MyData.MyDataParameter.MyDataType.DateTime)
			};
			DateTime StartTime = DateTime.Now;
			MyRecord.Say("1.0 准备开始...");
			using (MyData.MyDataTable mdata = new MyData.MyDataTable(SQL, mp))
			{
				var v = from a in mdata.MyRows
						select new Plan_ProduceNoteItem(a);
				MainData = v.ToList();
			}
			MyRecord.Say(string.Format("1.1 读取工单，耗时：{0}秒。", (DateTime.Now - StartTime).TotalSeconds));
			StartTime = DateTime.Now;
			///完工内容

			SQL = @"
SELECT [a].[ProduceNo] AS [ProdRdsNo], [a].[ProductCode] AS [ProdCode], (CASE WHEN [a].[PartID] = '' THEN '--' ELSE [a].[PartID] END) AS [PartID],
	   [mp].[id] AS [ProductID], [mpp].[reqnumb] AS [pNumb], [m].[name] AS [MachineName], [p].[name] AS [DepartmentName], [a].[Numb1] AS [FinishNumb],
	   [a].[Numb2] AS [RejectNumb], [a].[AdjustNumb] + [a].[SampleNumb] AS [LossedNumb], [a].[ProcessID] AS ProcessCode, [a].[CustCode],
	   [a].[StartTime] AS [BDD], [a].[EndTime] AS [EDD],[c].[Name] AS [ProdName]
  FROM [ProdDailyReport] [a]
	   INNER JOIN [_PMC_DeliverPlan_Sendlist] [b] ON [a].[ProduceNo] = [b].[PRODNO]
	   INNER JOIN [AllMaterialView] [c] ON [c].[code] = [a].[ProductCode]
	   INNER JOIN [moMachine] [m] ON [a].[MachinID] = [m].[code]
	   INNER JOIN [pbDept] [p] ON [m].[DepartmentID] = [p].[_ID]
	   INNER JOIN [moProduce] [mp] ON [mp].[rdsno] = [a].[ProduceNo]
	   INNER JOIN [moProdProcedure] [mpp] ON [mp].[id] = [mpp].[zbid]
											 AND [mpp].[id] = [a].[PrdID]
 WHERE [b].[SENDDATE] BETWEEN @ProdBegin AND @ProdEnd
";
			using (MyData.MyDataTable mdata = new MyData.MyDataTable(SQL, mp))
			{
				var v = from a in mdata.MyRows
						select new Plan_ProduceNoteItem(a);
				FinishData = v.ToList();
			}

			MyRecord.Say(string.Format("1.2 读取完工单，耗时：{0}秒。", (DateTime.Now - StartTime).TotalSeconds));
			StartTime = DateTime.Now;

			#region 备用
			/*
			///排程内容
			SQL = @"
Set NoCount On
Select Min(b.[_id]) as ID,a.Department,a.Process,b.ProdNo,b.PartNo,Bdd=Min(Bdd),Edd=Max(Edd)
Into #R
from _PMC_ProdPlan a,_PMC_ProdPlan_List b,_PMC_DeliverPlan_SendList c
Where a.[_id]=b.zbid And b.ProdID=c.ProdID And b.Side =0 And Not (b.Bdd is Null And b.Edd is Null) And ISNULL(b.YieldOff,0) = 0 And a.Status>=1 And
	  (a.Process = @Process or isNull(@Process,'')='') And
	  (a.Department =@DeptID or isNull(@DeptID,0)=0) And
	  (CharIndex('-' + @ProdType +'-',b.ProdCode)>0 or isNull(@ProdType,'')='') And
	  (b.ProdNo= @ProductRdsNo or isNull(@ProductRdsNo,'')='') And
	  (b.CustCode= @CustID or isNull(@CustID,'')='')
Group by b.ProdID,b.ProcID,a.Department,a.Process,b.ProdNo,b.PartNo

Select b.ProdNo as ProdRdsNo,b.ProdCode,PartID=Case When b.PartNo='' Then '--' Else b.PartNo End,
	   ProductID=b.ProdID,pNumb=b.ReqNumb,
	   ProcessCode=a.Process,b.CustCode,SendDate=a.Bdd
From #R a,_PMC_ProdPlan_List b
Where a.ID=b.[_ID]

Drop Table #R
Set NoCount OFF
";
			*/

			#endregion
			SQL = @"
SELECT [b].[PRODNO] AS [ProdRdsNo], [a].[ProdCode], (CASE WHEN [a].[PartNo] = '' THEN '--' ELSE [a].[PartNo] END) AS [PartID], [a].[DepartmentName],
	   [a].[MachineName], [b].[PRODID] AS [ProductID], [a].[PlanReqNumb] AS [pNumb], [a].[procno] AS [ProcessCode], [a].[CustCode], [a].[BDD], [a].[EDD],
	   [a].[Closer], [a].[CloseDate], [c].[Name] AS [ProdName], [a].[RdsNo] AS [PlanRdsNo], [b].[SENDDATE]
  FROM [_PMC_PlanProgressList_View] [a]
	   INNER JOIN [_PMC_DeliverPlan_Sendlist] [b] ON [a].zbid = [b].PRODID
	   INNER JOIN [AllMaterialView] [c] ON [a].ProdCode = [c].[code]
";

			using (MyData.MyDataTable mdata = new MyData.MyDataTable(SQL, mp))
			{
				var v = from a in mdata.MyRows
						select new Plan_ProduceNoteItem(a);
				PlanData = v.ToList();
			}

			MyRecord.Say(string.Format("1.3 读取分批交货，耗时：{0}秒。", (DateTime.Now - StartTime).TotalSeconds));
		}
		/// <summary>
		/// 加载内存中的表格结构
		/// </summary>
		private void LoadGridSchame(DateTime DateBegin, DateTime DateEnd)
		{
			Regex rgxMachineName = new Regex(@"[a-zA-Z0-9\u4e00-\u9fa5\#\/\~]*");
			string xMachineName = string.Empty;
			try
			{
				DateTime sDBegin = DateBegin, sDEnd = DateEnd;
				var v = from a in MainData
						orderby a.ProduceNote, a.SendDate, a.SortID, a.PartID
						where a.SendDate >= sDBegin && a.SendDate <= sDEnd
						group a by new { a.ProduceNote, a.PartID, a.SendDate } into g
						select new
						{
							g.Key.ProduceNote,
							g.Key.PartID,
							g.Key.SendDate,
							g.FirstOrDefault().ProdCode,
							g.FirstOrDefault().ProdName,
							g.FirstOrDefault().Inputer,
							g.FirstOrDefault().CustCode,
							g.FirstOrDefault().pNumb,
							g.FirstOrDefault().ProductNumb,
							g.FirstOrDefault().FinishRemark,
							g.FirstOrDefault().StatusName,
							g.FirstOrDefault().MachineName,
							g.FirstOrDefault().DepartmentName,
							g.FirstOrDefault().StockNumb
						};
				var v2 = from a in MainData
						 where a.SendDate >= sDBegin && a.SendDate <= sDEnd
						 group a by a.ProcessCode into g
						 orderby g.Key
						 select new
						 {
							 ProcessCode = g.Key,
							 ProcessName = g.FirstOrDefault().ProcessName
						 };

				if (v.Count() <= 0)
				{
					return;
				}

				GridList = new List<GridProduceNoteItem>();
				int uIndex = 0;
				foreach (var vRowItem in v)
				{
					uIndex++;
					MyRecord.Say(string.Format("计算第{0}条，共{1}条，{2}。", uIndex, v.Count(), vRowItem.ProduceNote));
					GridProduceNoteItem GridProdItem = new GridProduceNoteItem();
					GridList.Add(GridProdItem);
					GridProdItem.ID = uIndex;
					GridProdItem.ProduceRdsNo = vRowItem.ProduceNote;
					GridProdItem.ProduceNote = string.Format("{0}\r\n({1})\r\n{2}", vRowItem.ProduceNote, vRowItem.StatusName, vRowItem.FinishRemark);
					GridProdItem.ProductCode = vRowItem.ProdCode;
					GridProdItem.ProductName = vRowItem.ProdName;
					GridProdItem.ProductNumb = vRowItem.ProductNumb;
					GridProdItem.SendDate = vRowItem.SendDate;
					GridProdItem.pNumb = vRowItem.pNumb;
					GridProdItem.PartID = vRowItem.PartID;
					GridProdItem.StockNumb = vRowItem.StockNumb;
					GridProdItem.GridProcess = new List<GridProcessItem>();
					foreach (var vColItem in v2)
					{
						var vn = from a in MainData
								 where a.ProduceNote == vRowItem.ProduceNote && a.ProcessCode == vColItem.ProcessCode && a.PartID == vRowItem.PartID
								 select new
								 {
									 Numb1 = a.FinishNumb,
									 Numb2 = a.RejectNumb + a.LossedNumb,
									 Overdate = a.OverDate,
									 CloseDate = a.CloseDate,
									 Closer = a.Closer,
									 SRemark = a.SRemark,
									 DepartmentFullSortID = a.DepartmentFullSortID
								 };
						var vp = from a in PlanData
								 where a.ProduceNote == vRowItem.ProduceNote && a.ProcessCode == vColItem.ProcessCode && a.PartID == vRowItem.PartID
								 select new
								 {
									 Numb1 = a.pNumb,
									 BDD = a.BDD,
									 MachineName = a.MachineName,
									 DepartmentName = a.DepartmentName,
									 RdsNo = a.PlanRdsNo
								 };
						DateTime xPlanDateBegin = DateTime.Now.Hour > 20 ? DateTime.Now.Date.AddHours(19).AddMinutes(59) : DateTime.Now.Date.AddHours(7).AddMinutes(59);
						var vpc = from a in PlanData
								  where a.ProduceNote == vRowItem.ProduceNote && a.ProcessCode == vColItem.ProcessCode && a.PartID == vRowItem.PartID &&
										a.BDD > DateTime.MinValue &&
										((a.BDD - vRowItem.SendDate).TotalHours < 0 && (a.BDD - xPlanDateBegin).TotalHours > 0)
								  select new
								  {
									  Numb1 = a.pNumb,
									  BDD = a.BDD,
									  MachineName = a.MachineName,
									  DepartmentName = a.DepartmentName,
									  RdsNo = a.PlanRdsNo
								  };
						if (vn.Count() > 0)
						{
							GridProcessItem gpi = new GridProcessItem();
							GridProdItem.GridProcess.Add(gpi);
							gpi.TotalFinishNumb = vn.FirstOrDefault().Numb1;
							gpi.ProcessCode = vColItem.ProcessCode;
							gpi.ProcessName = vColItem.ProcessName;
							gpi.ProduceRdsNo = GridProdItem.ProduceRdsNo;
							gpi.ProductName = GridProdItem.ProductName;
							gpi.SRemark = vn.FirstOrDefault().SRemark;
							gpi.DepartmentFullSortID = vn.FirstOrDefault().DepartmentFullSortID;

							DateTime ov = vn.Max(a => a.Overdate);
							DateTime xDate = DateTime.MinValue;

							if (ov <= DateTime.MinValue)
							{
								DateTime cv = vn.Max(a => a.CloseDate);
								if (cv <= DateTime.MinValue)
								{
									if (vp.Count() > 0)
									{
										if (vpc.Count() <= 0)
										{
											gpi.PlanNumb = string.Format("{0:0}", vp.OrderBy(a => a.BDD).FirstOrDefault().Numb1);
											if (gpi.TotalFinishNumb > vRowItem.pNumb)
											{
												gpi.PlanNumb = gpi.PlanDate = gpi.PlanMachine = LCStr("完工数满足需求");
												gpi.GridStyle = GridProdItem.ErrorRowStyle = "Finished";
											}
											else
											{
												xDate = vp.Min(a => a.BDD);
												if (xDate > DateTime.Parse("2000-01-01"))
												{
													gpi.GridStyle = GridProdItem.ErrorRowStyle = "ErrorPlan";
													gpi.isError = GridProdItem.isError = true;
													DateTime xNowDate = DateTime.Now;
													DateTime xSendDate = vRowItem.SendDate; /// 工單的交期。
													DateTime sPlanBegin = xNowDate.Date.AddHours(8);
													if (xSendDate.Hour > 22) sPlanBegin = xNowDate.Date.AddHours(20);

													double hSend = (xDate - xSendDate).TotalHours, hPlanBegin = (xDate - sPlanBegin).TotalHours, hNow = (xDate - xNowDate).TotalHours;

													if (hSend > -2 && hSend < 0)
													{
														gpi.IssuesMemo = LCStr("交期2H以内");
														gpi.PlanDate = LCStr(string.Format("{0:MM/dd HH:mm}(交期2H以内)", xDate));
													}
													else if (hSend > 0)
													{
														gpi.IssuesMemo = LCStr("超出交期");
														gpi.PlanDate = LCStr(string.Format("{0:MM/dd HH:mm}(超出交期)", xDate));
													}
													else if (hPlanBegin < 0)
													{
														gpi.IssuesMemo = LCStr("当班未排");
														gpi.PlanDate = LCStr(string.Format("{0:MM/dd HH:mm}(当班未排)", xDate));
														gpi.GridStyle = GridProdItem.ErrorRowStyle = "NoPlan";
													}
													else if (hNow < -3)
													{
														gpi.IssuesMemo = LCStr("当班已排未结");
														gpi.PlanDate = LCStr(string.Format("{0:MM/dd HH:mm}(当班已排未结3H)", xDate));
													}
													else
													{
														gpi.IssuesMemo = LCStr("排程时间错误");
														gpi.IssuesMemo = gpi.PlanDate = LCStr(string.Format("{0:MM/dd HH:mm}(错误)", xDate));
													}

													xMachineName = vp.FirstOrDefault().MachineName;
													if (rgxMachineName.IsMatch(xMachineName)) xMachineName = rgxMachineName.Match(xMachineName).Value;
													gpi.PlanMachine = string.Format("{0}({1})", xMachineName, vp.FirstOrDefault().DepartmentName);
													gpi.DepartmentName = vp.FirstOrDefault().DepartmentName;

												}
												else
												{
													if (vColItem.ProcessCode != "9999")
													{
														gpi.isError = GridProdItem.isError = true;
														gpi.PlanDate = gpi.PlanNumb = LCStr("未排期");
														gpi.PlanMachine = string.Format("({0})", vp.FirstOrDefault().DepartmentName);
														gpi.GridStyle = GridProdItem.ErrorRowStyle = "NoPlan";
														gpi.DepartmentName = vp.FirstOrDefault().DepartmentName;
														gpi.IssuesMemo = LCStr("未排期");
													}
												}
											}
										}
										else
										{
											gpi.PlanNumb = string.Format("{0:0}", vpc.OrderBy(a => a.BDD).FirstOrDefault().Numb1);
											xDate = vpc.Min(a => a.BDD);
											gpi.PlanDate = string.Format("{0:MM/dd HH:mm}", xDate);
											xMachineName = vp.FirstOrDefault().MachineName;
											if (rgxMachineName.IsMatch(xMachineName)) xMachineName = rgxMachineName.Match(xMachineName).Value;
											gpi.PlanMachine = string.Format("({0})", xMachineName);
											gpi.DepartmentName = vp.FirstOrDefault().DepartmentName;
										}
									}
									else
									{
										if (vColItem.ProcessCode != "9999")
										{
											gpi.isError = GridProdItem.isError = true;
											gpi.PlanNumb = gpi.PlanDate = LCStr("未排期");
											gpi.PlanMachine = string.Format("({0})", vp.FirstOrDefault().DepartmentName);
											gpi.GridStyle = GridProdItem.ErrorRowStyle = "NoPlan";
											gpi.DepartmentName = vp.FirstOrDefault().DepartmentName;
											gpi.IssuesMemo = LCStr("未排期");
										}
									}
								}
								else
								{
									gpi.PlanNumb = LCStr("关闭");
									gpi.PlanDate = vn.Max(a => a.Closer);
									gpi.PlanMachine = cv.ToString("MM-dd HH:mm");
									gpi.GridStyle = GridProdItem.ErrorRowStyle = "Stoped";
								}
							}
							else
							{
								gpi.PlanNumb = gpi.PlanDate = gpi.PlanMachine = LCStr("已完");
								gpi.GridStyle = GridProdItem.ErrorRowStyle = "Finished";
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				MyRecord.Say(ex);
			}
		}
		/// <summary>
		/// 导出到Excel
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		void ExportToExcelDeliveryPlan(string FileName)
		{
			if (MainData.IsNotEmptySet() && GridList.IsNotEmptySet())
			{
				var vListGridRow = from a in GridList
								   where a.isError
								   select a;
				List<GridProcessItem> AllProcessItemList = new List<GridProcessItem>();
				foreach (var item in vListGridRow)
				{
					AllProcessItemList.AddRange(item.GridProcess);
				}

				var v2 = from a in AllProcessItemList
						 group a by a.ProcessCode into g
						 orderby g.Key
						 select new
						 {
							 ProcessCode = g.Key,
							 ProcessName = g.FirstOrDefault().ProcessName
						 };

				HSSFWorkbook wb = new HSSFWorkbook();
				HSSFSheet st = (HSSFSheet)wb.CreateSheet();
				st.DefaultColumnWidth = 10;
				HSSFCellStyle cellNomalStyle = (HSSFCellStyle)wb.CreateCellStyle(), cellIntStyle = (HSSFCellStyle)wb.CreateCellStyle(), cellDateStyle = (HSSFCellStyle)wb.CreateCellStyle();

				//HSSFPalette palette = wb.GetCustomPalette();

				cellNomalStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
				cellNomalStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
				cellNomalStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
				cellNomalStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
				cellNomalStyle.VerticalAlignment = VerticalAlignment.Center;
				cellNomalStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.General;
				cellNomalStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
				cellNomalStyle.WrapText = true;

				cellIntStyle.CloneStyleFrom(cellNomalStyle);
				cellIntStyle.DataFormat = wb.CreateDataFormat().GetFormat(@"#,###,###");

				cellDateStyle.CloneStyleFrom(cellNomalStyle);
				cellDateStyle.DataFormat = wb.CreateDataFormat().GetFormat(@"yyyy/MM/dd HH:mm");

				HSSFFont blodfont1 = (HSSFFont)wb.CreateFont();
				blodfont1.Boldweight = (short)FontBoldWeight.Bold;
				blodfont1.Color = HSSFColor.Turquoise.Index;

				HSSFFont blodfont2 = (HSSFFont)wb.CreateFont();
				blodfont2.Boldweight = (short)FontBoldWeight.Bold;
				blodfont2.Color = HSSFColor.DarkRed.Index;

				Dictionary<string, HSSFCellStyle> cellColl = new Dictionary<string, HSSFCellStyle>();

				HSSFCellStyle cds1 = (HSSFCellStyle)wb.CreateCellStyle();
				cds1.CloneStyleFrom(cellNomalStyle);
				cellColl.Add("Normal", cds1);

				HSSFCellStyle cds2 = (HSSFCellStyle)wb.CreateCellStyle();
				cds2.CloneStyleFrom(cellNomalStyle);
				cds2.FillPattern = FillPattern.SolidForeground;
				cds2.FillForegroundColor = HSSFColor.Rose.Index;
				cds2.SetFont(blodfont2);
				cellColl.Add("NoPlan", cds2);

				HSSFCellStyle cds3 = (HSSFCellStyle)wb.CreateCellStyle();
				cds3.CloneStyleFrom(cellNomalStyle);
				cds3.FillPattern = FillPattern.SolidForeground;
				cds3.FillForegroundColor = HSSFColor.Brown.Index;
				cds3.SetFont(blodfont1);
				cellColl.Add("ErrorPlan", cds3);

				HSSFCellStyle cds4 = (HSSFCellStyle)wb.CreateCellStyle();
				cds4.CloneStyleFrom(cellNomalStyle);
				cds4.FillPattern = FillPattern.SolidForeground;
				cds4.FillForegroundColor = HSSFColor.Grey40Percent.Index;
				cellColl.Add("Stoped", cds4);

				HSSFCellStyle cds5 = (HSSFCellStyle)wb.CreateCellStyle();
				cds5.CloneStyleFrom(cellNomalStyle);
				cds5.FillPattern = FillPattern.SolidForeground;
				cds5.FillForegroundColor = HSSFColor.LightCornflowerBlue.Index;
				cellColl.Add("Finished", cds5);



				HSSFRow rNoteCaption1 = (HSSFRow)st.CreateRow(1);
				Dictionary<string, int> colKey = new Dictionary<string, int>();

				for (int i = 0; i < titlewords.Length - 1; i++)
				{
					ICell xCellCaption = rNoteCaption1.CreateCell(i);
					xCellCaption.SetCellValue(LCStr(titlewords[i]));
					xCellCaption.CellStyle = cellNomalStyle;
					colKey.Add(titleKeys[i], i);
				}
				int j = 1, u = titlewords.Length - 1;

				foreach (var v2Item in v2)
				{
					ICell xCellCaption = rNoteCaption1.CreateCell(u);
					xCellCaption.SetCellValue(v2Item.ProcessName);
					xCellCaption.CellStyle = cellNomalStyle;
					colKey.Add(v2Item.ProcessCode, u);
					u++;
				}

				ICell xCellCaptionStockNumb = rNoteCaption1.CreateCell(u);
				xCellCaptionStockNumb.SetCellValue(LCStr(titlewords.LastOrDefault()));
				xCellCaptionStockNumb.CellStyle = cellNomalStyle;
				colKey.Add(titleKeys.LastOrDefault(), u);

				int uIndex = 0;
				foreach (var vitem in vListGridRow)
				{
					HSSFRow[] reRowItem = new HSSFRow[4];
					reRowItem[0] = (HSSFRow)st.CreateRow(j + 1);
					reRowItem[1] = (HSSFRow)st.CreateRow(j + 2);
					reRowItem[2] = (HSSFRow)st.CreateRow(j + 3);
					reRowItem[3] = (HSSFRow)st.CreateRow(j + 4);

					int iColIndex = colKey["ID"];
					for (int i = 0; i < 4; i++)
					{
						ICell fxCell1 = reRowItem[i].CreateCell(iColIndex);
						fxCell1.SetCellValue(vitem.ID);
						fxCell1.SetCellType(CellType.Numeric);
						fxCell1.CellStyle = cellIntStyle;
					}
					st.SetColumnWidth(iColIndex, 5 * 256);
					st.AddMergedRegion(new CellRangeAddress(j + 1, j + 4, iColIndex, iColIndex));

					iColIndex = colKey["ProductRdsNo"];
					for (int i = 0; i < 4; i++)
					{
						ICell fxCell2 = reRowItem[i].CreateCell(iColIndex);
						fxCell2.SetCellValue(vitem.ProduceNote);
						fxCell2.SetCellType(CellType.String);
						fxCell2.CellStyle = cellNomalStyle;
					}
					st.SetColumnWidth(iColIndex, 15 * 256);
					st.AddMergedRegion(new CellRangeAddress(j + 1, j + 4, iColIndex, iColIndex));

					iColIndex = colKey["ProductCode"];
					for (int i = 0; i < 4; i++)
					{
						ICell fxCell3 = reRowItem[i].CreateCell(iColIndex);
						fxCell3.SetCellValue(vitem.ProductCode);
						fxCell3.SetCellType(CellType.String);
						fxCell3.CellStyle = cellNomalStyle;
					}
					st.SetColumnWidth(iColIndex, 15 * 256);
					st.AddMergedRegion(new CellRangeAddress(j + 1, j + 4, iColIndex, iColIndex));

					iColIndex = colKey["ProductName"];
					for (int i = 0; i < 4; i++)
					{
						ICell fxCell4 = reRowItem[i].CreateCell(iColIndex);
						fxCell4.SetCellValue(vitem.ProductName);
						fxCell4.SetCellType(CellType.String);
						fxCell4.CellStyle = cellNomalStyle;
					}
					st.SetColumnWidth(iColIndex, 15 * 256);
					st.AddMergedRegion(new CellRangeAddress(j + 1, j + 4, iColIndex, iColIndex));

					iColIndex = colKey["ProductNumb"];
					for (int i = 0; i < 4; i++)
					{
						ICell fxCell5 = reRowItem[i].CreateCell(iColIndex);
						fxCell5.SetCellValue(vitem.ProductNumb);
						fxCell5.SetCellType(CellType.Numeric);
						fxCell5.CellStyle = cellIntStyle;
					}
					st.SetColumnWidth(iColIndex, 8 * 256);
					st.AddMergedRegion(new CellRangeAddress(j + 1, j + 4, iColIndex, iColIndex));

					iColIndex = colKey["SendDate"];
					for (int i = 0; i < 4; i++)
					{
						ICell fxCell6 = reRowItem[i].CreateCell(iColIndex);
						fxCell6.SetCellValue(vitem.SendDate);
						fxCell6.CellStyle = cellDateStyle;
					}
					st.SetColumnWidth(iColIndex, 15 * 256);
					st.AddMergedRegion(new CellRangeAddress(j + 1, j + 4, iColIndex, iColIndex));

					iColIndex = colKey["pNumb"];
					for (int i = 0; i < 4; i++)
					{
						ICell fxCell7 = reRowItem[i].CreateCell(iColIndex);
						fxCell7.SetCellValue(vitem.pNumb);
						fxCell7.SetCellType(CellType.Numeric);
						fxCell7.CellStyle = cellIntStyle;
					}
					st.SetColumnWidth(iColIndex, 8 * 256);
					st.AddMergedRegion(new CellRangeAddress(j + 1, j + 4, iColIndex, iColIndex));

					iColIndex = colKey["PartID"];
					for (int i = 0; i < 4; i++)
					{
						ICell fxCell8 = reRowItem[i].CreateCell(iColIndex);
						fxCell8.SetCellValue(vitem.PartID);
						fxCell8.SetCellType(CellType.String);
						fxCell8.CellStyle = cellNomalStyle;
					}
					st.SetColumnWidth(iColIndex, 8 * 256);
					st.AddMergedRegion(new CellRangeAddress(j + 1, j + 4, iColIndex, iColIndex));

					iColIndex = colKey["StockNumb"];
					for (int i = 0; i < 4; i++)
					{
						ICell fxCell9 = reRowItem[i].CreateCell(iColIndex);
						fxCell9.SetCellValue(vitem.StockNumb);
						fxCell9.SetCellType(CellType.Numeric);
						fxCell9.CellStyle = cellIntStyle;
					}
					st.SetColumnWidth(iColIndex, 8 * 256);
					st.AddMergedRegion(new CellRangeAddress(j + 1, j + 4, iColIndex, iColIndex));


					iColIndex = colKey["Title"];
					ICell fxCellTitle1 = reRowItem[0].CreateCell(iColIndex);
					fxCellTitle1.SetCellValue(LCStr("总完工数："));
					fxCellTitle1.CellStyle = cellNomalStyle;
					ICell fxCellTitle2 = reRowItem[1].CreateCell(iColIndex);
					fxCellTitle2.SetCellValue(LCStr("排期数："));
					fxCellTitle2.CellStyle = cellNomalStyle;
					ICell fxCellTitle3 = reRowItem[2].CreateCell(iColIndex);
					fxCellTitle3.SetCellValue(LCStr("排期时间："));
					fxCellTitle3.CellStyle = cellNomalStyle;
					ICell fxCellTitle4 = reRowItem[3].CreateCell(iColIndex);
					fxCellTitle4.SetCellValue(LCStr("排程机台："));
					fxCellTitle4.CellStyle = cellNomalStyle;
					st.SetColumnWidth(iColIndex, 12 * 256);

					u = titlewords.Length - 1;

					foreach (var v2Item in v2)
					{
						reRowItem[0].CreateCell(u).CellStyle = cellNomalStyle;
						reRowItem[1].CreateCell(u).CellStyle = cellNomalStyle;
						reRowItem[2].CreateCell(u).CellStyle = cellNomalStyle;
						reRowItem[3].CreateCell(u).CellStyle = cellNomalStyle;
						st.SetColumnWidth(u, 12 * 256);
						u++;
					}


					foreach (var gProcessItem in vitem.GridProcess)
					{
						iColIndex = colKey[gProcessItem.ProcessCode];
						ICell xCell1 = reRowItem[0].GetCell(iColIndex);
						xCell1.SetCellValue(string.Format("{0:#,###,###,###}", gProcessItem.TotalFinishNumb));
						xCell1.SetCellType(CellType.String);
						ICell xCell2 = reRowItem[1].GetCell(iColIndex);
						xCell2.SetCellValue(gProcessItem.PlanNumb);
						xCell2.SetCellType(CellType.String);
						ICell xCell3 = reRowItem[2].GetCell(iColIndex);
						xCell3.SetCellValue(gProcessItem.PlanDate);
						xCell3.SetCellType(CellType.String);
						ICell xCell4 = reRowItem[3].GetCell(iColIndex);
						xCell4.SetCellValue(gProcessItem.PlanMachine);
						xCell4.SetCellType(CellType.String);
						if (gProcessItem.GridStyle.IsNotEmpty())
						{
							xCell1.CellStyle = cellColl[gProcessItem.GridStyle];
							xCell2.CellStyle = cellColl[gProcessItem.GridStyle];
							xCell3.CellStyle = cellColl[gProcessItem.GridStyle];
							xCell4.CellStyle = cellColl[gProcessItem.GridStyle];
						}
						else
						{
							xCell1.CellStyle = cellNomalStyle;
							xCell2.CellStyle = cellNomalStyle;
							xCell3.CellStyle = cellNomalStyle;
							xCell4.CellStyle = cellNomalStyle;
						}
					}
					j += 4;
					uIndex++;
				}


				string filename = FileName;
				if (File.Exists(filename)) File.Delete(filename);
				using (FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write))
				{
					wb.Write(fs);
				}
			}

		}

		protected class Plan_ProduceNoteItem
		{
			public Plan_ProduceNoteItem(MyData.MyDataRow mmdr)
			{
				ProduceNote = mmdr.Value("ProdRdsNo");
				ProdCode = mmdr.Value("ProdCode");
				ProdName = mmdr.Value("ProdName");
				PartID = mmdr.Value("PartID");
				ProductID = mmdr.IntValue("ProductID");
				pNumb = mmdr.IntValue("pNumb");
				FinishNumb = mmdr.IntValue("FinishNumb");
				RejectNumb = mmdr.IntValue("RejectNumb");
				LossedNumb = mmdr.IntValue("LossedNumb");
				SendDate = mmdr.DateTimeValue("SendDate");
				Inputer = mmdr.Value("Inputer");
				CustCode = mmdr.Value("CustCode");
				ProcessCode = mmdr.Value("ProcessCode");
				ProcessName = mmdr.Value("ProcessName");
				ProductNumb = mmdr.IntValue("ProductNumb");
				SortID = mmdr.IntValue("SortID");
				OverDate = mmdr.DateTimeValue("OverDate");
				StatusName = mmdr.Value("SName");
				FinishRemark = mmdr.Value("FinishRemark");
				Closer = mmdr.Value("Closer");
				CloseDate = mmdr.DateTimeValue("CloseDate");
				DepartmentName = mmdr.Value("DepartmentName");
				MachineName = mmdr.Value("MachineName");
				PlanRdsNo = mmdr.Value("PlanRdsNo");
				BDD = mmdr.DateTimeValue("BDD");
				Edd = mmdr.DateTimeValue("EDD");
				SRemark = mmdr.Value("SRemark");
				DepartmentFullSortID = mmdr.Value("DepartmentFullSortID");
			}

			/// <summary>
			/// 部件
			/// </summary>
			public string PartID { get; private set; }

			/// <summary>
			/// 产品编号
			/// </summary>
			public string ProdCode { get; private set; }

			/// <summary>
			/// 产品名称
			/// </summary>
			public string ProdName { get; private set; }

			/// <summary>
			/// 工单号
			/// </summary>
			public string ProduceNote { get; private set; }

			/// <summary>
			/// 客户编号
			/// </summary>
			public string CustCode { get; private set; }

			/// <summary>
			/// 工序编号
			/// </summary>
			public string ProcessCode { get; private set; }

			/// <summary>
			/// 工序名称
			/// </summary>
			public string ProcessName { get; private set; }

			/// <summary>
			/// 刻交期人
			/// </summary>
			public string Inputer { get; private set; }

			/// <summary>
			/// 出货明细数量
			/// </summary>
			public int pNumb { get; private set; }

			/// <summary>
			/// 完工正品数量
			/// </summary>
			public int FinishNumb { get; private set; }

			/// <summary>
			/// 完工不良品数
			/// </summary>
			public int RejectNumb { get; private set; }

			/// <summary>
			/// 完工损耗数(过版纸)
			/// </summary>
			public int LossedNumb { get; private set; }

			/// <summary>
			/// 交期
			/// </summary>
			public DateTime SendDate { get; private set; }

			/// <summary>
			/// 上机日期
			/// </summary>
			public DateTime BDD { get; private set; }
			/// <summary>
			/// 下机日期
			/// </summary>
			public DateTime Edd { get; private set; }

			/// <summary>
			/// 完成日期
			/// </summary>
			public DateTime OverDate { get; private set; }

			/// <summary>
			/// 工单ID
			/// </summary>
			public int ProductID { get; private set; }

			/// <summary>
			/// 订单数
			/// </summary>
			public int ProductNumb { get; private set; }

			/// <summary>
			/// 排序
			/// </summary>
			public int SortID { get; private set; }

			public string FinishRemark { get; private set; }

			public string StatusName { get; private set; }

			public string Closer { get; private set; }

			public DateTime CloseDate { get; private set; }

			public string DepartmentName { get; private set; }

			public string MachineName { get; private set; }
			public double StockNumb { get; set; }
			public string PlanRdsNo { get; private set; }

			public string SRemark { get; private set; }
			public string DepartmentFullSortID { get; private set; }
		}

		public class GridProduceNoteItem
		{
			/// <summary>
			/// 序号
			/// </summary>
			public int ID { get; set; }
			public string ProduceRdsNo { get; set; }
			/// <summary>
			/// 工单号
			/// </summary>
			public string ProduceNote { get; set; }
			/// <summary>
			/// 产品号
			/// </summary>
			public string ProductCode { get; set; }
			/// <summary>
			/// 料号
			/// </summary>
			public string ProductName { get; set; }
			/// <summary>
			/// 订单数
			/// </summary>
			public double pNumb { get; set; }
			/// <summary>
			/// 交期
			/// </summary>
			public DateTime SendDate { get; set; }
			/// <summary>
			/// 需求数
			/// </summary>
			public double ProductNumb { get; set; }
			/// <summary>
			/// 台号/部件编号
			/// </summary>
			public string PartID { get; set; }
			/// <summary>
			/// 库存数
			/// </summary>
			public double StockNumb { get; set; }

			public string StatusName { get; set; }

			public string MachineName { get; set; }

			public string DepartmentName { get; set; }
			/// <summary>
			/// 内容
			/// </summary>
			public string Remark { get; set; }
			/// <summary>
			/// 有问题
			/// </summary>
			public bool isError { get; set; }

			public string ErrorRowStyle { get; set; }
			/// <summary>
			/// 工序
			/// </summary>
			public List<GridProcessItem> GridProcess { get; set; }

		}

		public class GridProcessItem
		{
			public string ProcessName { get; set; }

			public string ProcessCode { get; set; }

			public double TotalFinishNumb { get; set; }

			public string PlanNumb { get; set; }

			public string PlanDate { get; set; }

			public string PlanMachine { get; set; }

			public string GridStyle { get; set; }

			public string DepartmentName { get; set; }

			public string DepartmentFullSortID { get; set; }

			public bool isError { get; set; }

			public string IssuesMemo { get; set; }
			public string ProduceRdsNo { get; set; }
			public string ProductName { get; set; }

			public string SRemark { get; set; }
		}


		#endregion

	}
}
