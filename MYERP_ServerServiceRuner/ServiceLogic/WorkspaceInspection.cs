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

        #region 发送3天前没有结案的工单

        void SendWorkspaceInspectionLoder()
        {
            MyRecord.Say("开启定时发送当日纪律稽核表..........");
            Thread t = new Thread(SendWorkspaceInspectionEmail);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启定时发送当日纪律稽核表完成。");
        }

        void SendWorkspaceInspectionEmail()
        {
            try
            {
                MyRecord.Say("-----------------开启定时发送当日纪律稽核表-------------------------");
                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{3}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容为昨天（{0:yyyy年MM月dd日}）品保纪律稽核问题反馈。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (各部门纪律得分请看附档。)</FONT></DIV>
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
    白夜班
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    部门
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    工序
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    机台
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    缺失类别
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    纪律稽核不佳说明
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    责任人
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    责任主管
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    稽核人
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    稽核日期
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
    </TR>
";

                string SQL = @"
Select a.InspectDate,a.ClassType,
       ClassTypeName =(Case When isNull(a.ClassType,0) = 1 Then '白班' Else '夜班' End),
	   DeptName = dp.name,dp.FullSortID,
	   ProcessName = (Select Name from moProcedure pp where pp.code = b.ProcessCode),
	   MachineName = (select Name From moMachine mm where mm.code = b.MachineCode),
	   b.MachineCode,b.ProcessCode,cm.ProjName,
	   a.InspectMan,b.InspectMemo,b.Remark2,b.Remark3
  from [_PMC_KanbanKPI_CheckList] a Inner Join [_PMC_KanbanKPI_CheckList_List] b On a.[_id] = b.zbid
                                    Inner Join [pbDept] dp On dp.Code = b.DepartmentCode
                                    Inner Join [_PMC_Machine_CheckItemList] cm On cm.ItemCode = b.ItemCode
Where a.InspectDate = @DateBegin And b.ScoreSelectValue = 3
Order by a.InspectDate,a.ClassType,dp.FullSortID,b.ProcessCode,b.MachineCode
";
                DateTime NowTime = DateTime.Now;
                string brs = "";
                DateTime dt1 = NowTime.DayOfWeek == DayOfWeek.Monday ? NowTime.AddDays(-2).Date : NowTime.AddDays(-1).Date;
                MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
                {
                    new MyData.MyDataParameter("@DateBegin",dt1, MyData.MyDataParameter.MyDataType.DateTime)
                };
                MyRecord.Say(string.Format("后台计算定时发送当日纪律稽核表，{0}", dt1));

                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, mps))
                {
                    string fname = string.Empty;
                    MyRecord.Say("创建SendMail。");
                    MyBase.SendMail sm = new MyBase.SendMail();
                    MyRecord.Say("加载邮件内容。");
                    MyRecord.Say("计算定时发送当日纪律稽核表");
                    if (md.MyRows.IsNotEmptySet())
                    {

                        foreach (var ri in md.MyRows)
                        {
                            brs += string.Format(br, ri.Value("ClassTypeName"),
                                                     ri.Value("DeptName"),
                                                     ri.Value("ProcessName"),
                                                     ri.Value("MachineName"),
                                                     ri.Value("ProjName"),
                                                     ri.Value("InspectMemo"),
                                                     ri.Value("Remark2"),
                                                     ri.Value("Remark3"),
                                                     ri.Value("InspectMan"),
                                                     string.Format("{0:yyyy/MM/dd}", ri.DateTimeValue("InspectDate"))
                                                    );
                        }
                        MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", md.Rows.Count));
                        bbr = string.Format(bbr, brs);
                    }
                    else
                    {
                        bbr = @"<FONT size=5 face=PMingLiU color=Red>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 没有发现纪律稽核不佳项。</FONT></DIV>";
                    }
                    sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, dt1, bbr, NowTime, MyBase.CompanyTitle, LocalInfo.GetLocalIp()));
                    sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_纪律稽核结果。", NowTime, MyBase.CompanyTitle));

                    string SQL2 = @"
Select a.InspectDate,a.ClassType,
       ClassTypeName =(Case When isNull(a.ClassType,0) = 1 Then '白班' Else '夜班' End),
	   Score=(Sum(b.Score) / sum(cm.Result1Value)) * 100,DeptName = Max(dp.Name),dp.FullSortID,
	   Value3Time = Sum(Case When b.ScoreSelectValue = 3 Then 1 Else 0 End)
  from [_PMC_KanbanKPI_CheckList] a Inner Join [_PMC_KanbanKPI_CheckList_List] b On a.[_id] = b.zbid
                                    Inner Join [pbDept] dp On dp.Code = b.DepartmentCode
                                    Inner Join [_PMC_Machine_CheckItemList] cm On cm.ItemCode = b.ItemCode
Where a.InspectDate = @DateBegin
Group by dp.FullSortID,b.DepartmentCode,a.InspectDate,a.ClassType
Order by a.InspectDate,a.ClassType,dp.FullSortID
";
                    MyRecord.Say("计算邮件附件内容。");
                    using (MyData.MyDataTable md2 = new MyData.MyDataTable(SQL2, mps))
                    {
                        var vAttachView = from a in md2.MyRows
                                          orderby a.Value("ClassType"), a.Value("FullSortID")
                                          select a;
                        string[] fields = new string[] { "ClassTypeName", "DeptName", "Score", "Value3Time" };
                        string[] captions = new string[] { "白夜班", "部门", "得分", "不佳数" };
                        fname = ExportExcel.Export(sm, vAttachView, fields, captions, LCStr("纪律稽核得分"));
                    };

                    //string mailto = ConfigurationManager.AppSettings["WorkspaceInspectionMailTo"], mailcc = ConfigurationManager.AppSettings["WorkspaceInspectionMailCC"];
                    //MyRecord.Say(string.Format("MailTO:{0}    MailCC:{1}", mailto, mailcc));
                    //sm.MailTo = mailto; // "my18@my.imedia.com.tw,xang@my.imedia.com.tw,lghua@my.imedia.com.tw,my64@my.imedia.com.tw";
                    //sm.MailCC = mailcc; // "jane123@my.imedia.com.tw,lwy@my.imedia.com.tw,my80@my.imedia.com.tw";
                    ////sm.MailTo = "my80@my.imedia.com.tw";

                    MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("WorkspaceInspection");
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

        #endregion


    }
}
