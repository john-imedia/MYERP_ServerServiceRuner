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


		#region 發送每日出货计划数量安排
		void ECRECNProgressSender()
		{
			MyRecord.Say(@"开启发送每日未结ECR\ECN..........");
			Thread t = new Thread(ECRECNProgressSendMail);
			t.IsBackground = true;
			t.Start();
			MyRecord.Say(@"定时发送每日未结ECR、ECN已经启动。");
		}


		void ECRECNProgressSendMail()
		{
			try
			{
				MyRecord.Say("------------------开始发送ECRECN跟踪表----------------------------");
				MyRecord.Say("从数据库搜寻内容");
				DateTime NowTime = DateTime.Now;
				string body = @"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#e6f3af #ffffff>
<DIV><FONT size=3 face=PMingLiU>{0}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;</FONT></DIV>
<DIV><FONT color=#ff0000 size=5 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <B>{2:yy/MM/dd} {3} </B></FONT></DIV>
{4}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=2 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，切勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{1:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
";
				string brs = MyConvert.ZH_TW(@"
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	发起部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	客户
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	料号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	ECR单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	建立
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	状态
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	IE
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	SR
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	生控
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	业务
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工程
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	核准
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	ECN单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	ECN确认
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	ECN状态
	</TD>
	</TR>
	{0}
</TBODY></TABLE></FONT>
</DIV>
");
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
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{5}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{6}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{7}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{8}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{9}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{10}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{11}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{12}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{13}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{14}
	</TD>
	</TR>
";

				string mailBodyMainSentence = string.Empty;
				MyRecord.Say("读取数据");
				string SQL = @"
SELECT [a].[RdsNo] AS [ECR], [a].[InputDate] AS [ECRInputDate], [a].[RequestDepartmentName], [a].[RequestAuthor] AS [ECRAuthor], 
	   [a].[RequestConfirmAuthor] AS [ECRConfirmAuthor], [a].[RequestDate] AS [ECRDate], 
	   [a].[CheckDate] AS [ECRCheckDate],[a].[Checker] AS [ECRChecker],
	   [a].[ReviewAuthor1], [a].[ReviewAuthor2], [a].[ReviewAuthor3], [a].[ReviewAuthor4_1] AS [ReviewAuthor4], [a].[ReviewAuthor5], [a].[ReviewAuthor6],
	   [a].[ReviewAuthor7], [a].[ReviewAuthor8], [a].[ReviewDate1], [a].[ReviewDate2], [a].[ReviewDate3], [a].[ReviewDate4_1] AS [ReviewDate4], [a].[ReviewDate5],
	   [a].[ReviewDate6], [a].[ReviewDate7], [a].[ReviewDate8], (SELECT [Name]
																   FROM [_SY_Status] [RST]
																  WHERE [RST].[Type] = 'ECR'
																		AND [RST].[StatusID] = [a].[Status]) AS [ECRStatusName], [b].[RdsNo] AS [ECN],
	   [b].[RequestConfirmAuthor] AS [ECNConfirmAuthor], b.[RequestConfirmDate] AS [ECNConfirmDate], b.[CheckDate] AS [ECNCheckDate],b.Inputer as [ECNAuthor],
	   b.[Checker] AS [ECNChecker], (SELECT [Name]
									   FROM [_SY_Status] [RST]
									  WHERE [RST].[Type] = 'ECN'
											AND [RST].[StatusID] = [b].[Status]) AS [ECNStatusName],
	   [a].[Code],[a].[MutiProd],[a].[ProdName],[a].[CustCode]
  FROM [_PMC_EngineerChangeRequest] [a]
	   LEFT OUTER JOIN [_PMC_EngineerChangeNote] [b] ON [a].[ECN] = [b].[RdsNo]
 WHERE [b].[CheckDate] IS NULL
	   AND [a].[Status] > -1
	   AND a.[InputDate] > '2018-06-01 00:00:00'
 ORDER BY [a].[RdsNo];   
";
				MyRecord.Say("读取计算完成。");
				MyRecord.Say("创建SendMail。");
				MyBase.SendMail sm = new MyBase.SendMail();
				string fname = string.Empty;
				using (MyData.MyDataTable md = new MyData.MyDataTable(SQL))
				{
					if (md.MyRows.IsNotEmptySet())
					{
						string xbr = string.Empty;
						foreach (var r in md.MyRows)
						{
							xbr += string.Format(br, r.Value("RequestDepartmentName"),
													 r.Value("CustCode"),
													 r.BooleanValue("MutiProd") ? r.Value("Code") : r.Value("ProdName"),
													 r.Value("ECR"),
													 string.Format("{0}{2}({1:yy/MM/dd HH:mm})", r.Value("ECRAuthor"), r.DateTimeValue("ECRInputDate"), r.Value("ECRConfirmAuthor").IsNotEmpty() ? string.Format(@"/{0}", r.Value("ECRConfirmAuthor")) : ""),
													 r.Value("ECRStatusName"),
													 r.DateTimeValue("ReviewDate1").IsNotEmpty() ? string.Format("{0}({1:yy/MM/dd HH:mm})", r.Value("ReviewAuthor1"), r.DateTimeValue("ReviewDate1")) : "",
													 r.DateTimeValue("ReviewDate2").IsNotEmpty() ? string.Format("{0}({1:yy/MM/dd HH:mm})", r.Value("ReviewAuthor2"), r.DateTimeValue("ReviewDate2")) : "",
													 r.DateTimeValue("ReviewDate3").IsNotEmpty() ? string.Format("{0}({1:yy/MM/dd HH:mm})", r.Value("ReviewAuthor3"), r.DateTimeValue("ReviewDate3")) : "",
													 r.DateTimeValue("ReviewDate5").IsNotEmpty() ? string.Format("{0}({1:yy/MM/dd HH:mm})", r.Value("ReviewAuthor5"), r.DateTimeValue("ReviewDate5")) : "",
													 r.DateTimeValue("ReviewDate6").IsNotEmpty() ? string.Format("{0}({1:yy/MM/dd HH:mm})", r.Value("ReviewAuthor6"), r.DateTimeValue("ReviewDate6")) : "",
													 r.DateTimeValue("ReviewDate8").IsNotEmpty() ? string.Format("{0}({1:yy/MM/dd HH:mm})", r.Value("ReviewAuthor8"), r.DateTimeValue("ReviewDate8")) : "",
													 r.Value("ECN"),
													 r.DateTimeValue("ECNConfirmDate").IsNotEmpty() ? string.Format("{0}({1:yy/MM/dd HH:mm})", r.Value("ECNConfirmAuthor"), r.DateTimeValue("ECNConfirmDate")) : "",
													 r.Value("ECNStatusName")
													 );
						}
						brs = string.Format(brs, xbr);
						MyRecord.Say("生成邮件内容。");
						mailBodyMainSentence = "未结案ECR/ECN单进度。";
					}
					else
					{
						mailBodyMainSentence = "无未结案ECR/ECN单。";
						brs = string.Empty;
					}
				}

				MyRecord.Say("加载邮件内容。");

				sm.MailBodyText = string.Format(body, MyBase.CompanyTitle, DateTime.Now, NowTime, mailBodyMainSentence, brs);
				sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}未结ECR/ECN进度表", NowTime.Date, MyBase.CompanyTitle));

				MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("ECRECN");
				MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
				sm.MailTo = mAddress.MailTo;
				sm.MailCC = mAddress.MailCC;
				//sm.MailTo = "my80@my.imedia.com.tw";
				MyRecord.Say("发送邮件。");
				sm.SendOut();
				sm.mail.Dispose();
				sm = null;
				MyRecord.Say("已经发送。");
				MyRecord.Say(@"------------------每日未结ECR\ECN-发送完成----------------------------");
			}
			catch (Exception ex)
			{
				MyRecord.Say(ex);
			}
		}

		#endregion
	}
}
