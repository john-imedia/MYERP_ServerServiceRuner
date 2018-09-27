using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;
using System.Web;
using Microsoft.ClearScript.V8;
using Microsoft.ClearScript;
using Microsoft.ClearScript.Windows;
using MYERP_ServerServiceRuner.Base;

namespace MYERP_ServerServiceRuner
{
	partial class MainService
	{

		#region 发送不合格判定责任单位统计表

		void SendIPQCListStaticLoader()
		{
			MyRecord.Say("开启定时发送4H不良判定统计表..........");
			Thread t = new Thread(SendIPQCListStatic);
			t.IsBackground = true;
			t.Start();
			MyRecord.Say("开启定时发送4H不良判定统计表完成。");
		}

		void SendIPQCListStatic()
		{
			try
			{
				MyRecord.Say("-----------------开启定时发送4H不良判定统计表-------------------------");
				string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{3}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 下述表格为过去4H（{0:MM.dd HH:mm}至{2:MM.dd HH:mm}）制程不良判定结果的统计。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (详细不良判定内容，请看附档。)</FONT></DIV>
<DIV>{1}
</DIV>
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{2:yy/MM/dd HH:mm}，由ERP系统伺服器（{4}）自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");
				string bbr = @"
<FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	责任部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	前5大不良
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	平均不良率
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	不良数最高工序
	</TD>
	</TR>
	{0}
</TBODY></TABLE>
</FONT>
";
				string br = @"
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{0}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{1}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{2:#.###%}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{3}
	</TD>
	</TR>
";

				string SQL = @"
	SELECT [a].[ProcessID] AS [OccuredProcessCode], [a].[MachinID] AS [OccuredMachineCode], [a].[Numb1], [a].[Numb2], [a].[SampleNumb], [a].[AdjustNumb],
			[a].[AccNoteRdsNo], [a].[RenoteRdsno], [a].[ProduceNo] AS [ProduceRdsNo], [a].[ProductCode], [b].[RejNumb], [b].[ScrapNumber],
			[b].[AssignToDepartmentID], [b].[ProjCode], [b].[ItemID], [d].[name] AS [ProductName], [b].[Editor], [b].[EditDate], [c].[name] AS [ProjName],
			[dp].[name] AS [AssignToDepartmentName], [dp].[FullSortID], [p].[name] AS [AssignToProcessName], [p1].[name] AS [OccuredProcessName],
			[m].[name] AS [OccuredMachineName], [dp1].[name] AS [OccuredDepartmentName], [a].[RptDate],
			(ISNULL([b].[RejNumb], 0) / (ISNULL([a].[Numb1], 0) + ISNULL([a].[Numb2], 0) + ISNULL([a].[SampleNumb], 0) + ISNULL([a].[AdjustNumb], 0))) AS [Yield]
		FROM [dbo].[ProdDailyReport] [a]
			INNER JOIN [dbo].[_PMC_IPQC_List] [b] ON [a].[_ID] = [b].[zbid]
			INNER JOIN [dbo].[AllMaterialView] [d] ON [d].[code] = [a].[ProductCode]
			INNER JOIN [dbo].[moProcedure] [p1] ON [p1].[code] = [a].[ProcessID]
			INNER JOIN [dbo].[moMachine] [m] ON [m].[code] = [a].[MachinID]
			INNER JOIN [dbo].[pbDept] [dp1] ON [dp1].[_ID] = [m].[DepartmentID]
			LEFT OUTER JOIN [dbo].[_QC_Item] [c] ON [b].[ProjCode] = [c].[code]
			LEFT OUTER JOIN [dbo].[pbDept] [dp] ON [dp].[_ID] = [b].[AssignToDepartmentID]
			LEFT OUTER JOIN [dbo].[moProcedure] [p] ON [p].[code] = [b].[AssignToProcessCode]
		WHERE [a].[QCCheckDate] BETWEEN DATEADD(HOUR, -4, GETDATE()) AND GETDATE()
		ORDER BY [a].[RptDate], [a].[ProduceNo], [a].[AccNoteRdsNo]
";
				DateTime NowTime = DateTime.Now;
				string brs = "";
				MyRecord.Say(string.Format("后台计算定时4H不良判定统计表，{0}", NowTime));
				using (MyData.MyDataTable md = new MyData.MyDataTable(SQL))
				{
					string fname = string.Empty;
					MyRecord.Say("创建SendMail。");
					MyBase.SendMail sm = new MyBase.SendMail();
					MyRecord.Say("加载邮件内容。");
					MyRecord.Say("计算定时发送4H不良判定统计表");
					if (md.MyRows.IsNotEmptySet())
					{
						var v = from a in md.MyRows
								select new IPQCStaticItem(a);
						List<IPQCStaticItem> lData = v.ToList();
						var va = from a in lData
								 group a by a.AssignToDepartmentName into g
								 select new
								 {
									 DepartmentName = g.Key,
									 Top5 = string.Join("，", g.GroupBy(k => k.ProjName).Select(q => new { ProjName = q.Key, RejNumb = q.Sum(p => p.RejNumb) }).OrderByDescending(x => x.RejNumb).Take(5).Select(y => string.Format("{0}：{1}", y.ProjName, y.RejNumb))),
									 AvgYeild = g.Average(z => z.Yield),
									 TopProcessName = g.OrderBy(u => u.RejNumb).Take(1).FirstOrDefault().AssignToProcessName
								 };
						foreach (var ri in va)
						{
							brs += string.Format(br, ri.DepartmentName, ri.Top5, ri.AvgYeild, ri.TopProcessName);
						}
						MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", md.Rows.Count));
						bbr = string.Format(bbr, brs);
					}
					else
					{
						bbr = @"<FONT size=5 face=PMingLiU color=Red>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 过去4H没有做过不良判定。</FONT></DIV>";
					}

					sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, NowTime.AddHours(-4), bbr, NowTime, MyBase.CompanyTitle, LocalInfo.GetLocalIp()));
					sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日HH时}_4H不良判定统计。", NowTime, MyBase.CompanyTitle));

					MyRecord.Say("计算邮件附件内容。");
					if (md.MyRows.IsNotEmptySet())
					{
						IPQCStaticItem xx;
						var vAttachView = from a in md.MyRows
										  orderby a.Value("FullSortID")
										  select a;
						string[] fields = new string[] { "ProduceRdsNo", "AccNoteRdsNo", "RptDate", "OccuredDepartmentName", "OccuredProcessName", "OccuredMachineName", "RenoteRdsno", "ProductCode", "ProductName", "Numb1", "Numb2", "SampleNumb", "AdjustNumb", "ProjName", "RejNumb", "Editor", "EditDate", "AssignToDepartmentName", "AssignToProcessName", "Yield" };
						string[] captions = new string[] { "工单号", "完工单号", "完工日期", "发现部门", "发现工序", "发现机台", "不不良判定单号", "产品编号", "料号", "良品数", "不良数", "样品数", "过版纸数", "不良项目", "不良数", "判定人", "判定时间", "责任部门", "责任机台", "不良率" };
						fname = ExportExcel.Export(sm, vAttachView, fields, captions, LCStr("不良判定表"));
					};

					MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("IPQCListStatic");
					MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
					sm.MailTo = mAddress.MailTo;
					sm.MailCC = mAddress.MailCC;
					//sm.MailTo = "my80@my.imedia.com.tw";

					MyRecord.Say("发送邮件。");
					sm.SendOut();
					sm.mail.Dispose();
					sm = null;
					MyRecord.Say("邮件已经发送。");
					if (File.Exists(fname))
					{
						File.Delete(fname);
					}
				}
				MyRecord.Say("------------------发送完成----------------------------");
			}
			catch (Exception e)
			{
				MyRecord.Say(e);
			}
		}


		class IPQCStaticItem
		{
			public IPQCStaticItem(MyData.MyDataRow r)
			{
				OccuredProcessCode = r.Value("OccuredProcessCode");
				OccuredMachineCode = r.Value("OccuredMachineCode");
				Numb1 = r.DoubleValue("Numb1");
				Numb2 = r.DoubleValue("Numb2");
				SampleNumb = r.DoubleValue("SampleNumb");
				AdjustNumb = r.DoubleValue("AdjustNumb");
				AccNoteRdsNo = r.Value("AccNoteRdsNo");
				RenoteRdsno = r.Value("RenoteRdsno");
				ProduceRdsNo = r.Value("ProduceRdsNo");
				ProductCode = r.Value("ProductCode");
				RejNumb = r.DoubleValue("RejNumb");
				ScrapNumber = r.DoubleValue("ScrapNumber");
				AssignToDepartmentID = r.IntValue("AssignToDepartmentID");
				ProjCode = r.Value("ProjCode");
				ItemID = r.IntValue("ItemID");
				ProductName = r.Value("ProductName");
				Editor = r.Value("Editor");
				EditDate = r.DateTimeValue("EditDate");
				ProjName = r.Value("ProjName");
				AssignToDepartmentName = r.Value("AssignToDepartmentName");
				FullSortID = r.Value("FullSortID");
				AssignToProcessName = r.Value("AssignToProcessName");
				OccuredProcessName = r.Value("OccuredProcessName");
				OccuredMachineName = r.Value("OccuredMachineName");
				OccuredDepartmentName = r.Value("OccuredDepartmentName");
				RptDate = r.DateTimeValue("RptDate");
				Yield = r.DoubleValue("Yield");
			}
			public string OccuredProcessCode { get; set; }
			public string OccuredMachineCode { get; set; }
			public double Numb1 { get; set; }
			public double Numb2 { get; set; }
			public double SampleNumb { get; set; }
			public double AdjustNumb { get; set; }
			public string AccNoteRdsNo { get; set; }
			public string RenoteRdsno { get; set; }
			public string ProduceRdsNo { get; set; }
			public string ProductCode { get; set; }
			public double RejNumb { get; set; }
			public double ScrapNumber { get; set; }
			public int AssignToDepartmentID { get; set; }
			public string ProjCode { get; set; }
			public int ItemID { get; set; }
			public string ProductName { get; set; }
			public string Editor { get; set; }
			public DateTime EditDate { get; set; }
			public string ProjName { get; set; }
			public string AssignToDepartmentName { get; set; }
			public string FullSortID { get; set; }
			public string AssignToProcessName { get; set; }
			public string OccuredProcessName { get; set; }
			public string OccuredMachineName { get; set; }
			public string OccuredDepartmentName { get; set; }
			public DateTime RptDate { get; set; }
			public double Yield { get; set; }
		}

		#endregion


	}
}
