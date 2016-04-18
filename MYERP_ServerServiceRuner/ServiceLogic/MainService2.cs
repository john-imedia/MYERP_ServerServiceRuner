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
using C1.Win.C1FlexGrid;
using System.Windows.Forms;
using System.Drawing;
using System.IO;


namespace MYERP_ServerServiceRuner
{
    public partial class MainService
    {
        #region 审核昨天的纪律单

        void WorkspaceInspectCheckLoader()
        {
            MyRecord.Say("开启库存后台计算..........");
            Thread t = new Thread(new ThreadStart(WorkspaceInspectChecker));
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启库存后台计算成功。");
        }

        void WorkspaceInspectChecker()
        {
            DateTime rxtime = DateTime.Now;
            string SQL = "";
            MyRecord.Say("---------------------启动定时审核纪律单。------------------------------");
            DateTime NowTime = DateTime.Now;
            DateTime StopTime = NowTime.Date.AddDays(-1);
            int tpid = (NowTime.Hour < 13) ? 1 : 2;
            MyRecord.Say(string.Format("审核截止时间：{0:yy/MM/dd HH:mm},ID={1}", StopTime, tpid));
            #region 审核纪律单
            try
            {
                MyRecord.Say("审核纪律单");
                SQL = @"Select * from _PMC_KanbanKPI_CheckList Where isNull(Status,0)=0 And ((InspectDate < @InputEnd) or (Convert(DateTime,Convert(VarChar(10),InspectDate,121)) = @InputEnd And ClassType=@TpID))";
                MyData.MyDataTable mTableWorkspaceInspect = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime), new MyData.MyParameter("@TpID", tpid, MyData.MyParameter.MyDataType.Int));
                if (_StopAll) return;
                if (mTableWorkspaceInspect != null && mTableWorkspaceInspect.MyRows.Count > 0)
                {
                    string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableWorkspaceInspect.MyRows.Count);
                    MyRecord.Say(memo);
                    int mTableWorkspaceInspectCount = 1;
                    foreach (MyData.MyDataRow r in mTableWorkspaceInspect.MyRows)
                    {
                        if (_StopAll) return;
                        int CurID = Convert.ToInt32(r["_ID"]);
                        string RdsNo = Convert.ToString(r["RdsNo"]);
                        memo = string.Format("审核第{0}条，纪律检查单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableWorkspaceInspectCount, RdsNo, CurID, r["InputDate"]);
                        MyRecord.Say(memo);
                        if (!RecordCheck("_PMC_KanbanKPI_CheckList", "_PMC_KanbanKPI_CheckList_StatusRecorder", CurID, RdsNo))
                        {
                            MyRecord.Say("审核错误！");
                        }

                        MyData.MyParameter mdpID = new MyData.MyParameter("@id", CurID, MyData.MyParameter.MyDataType.Int);
                        string SQLDelete = "Delete from _PMC_KanbanKPI_CheckList_List Where isNull(Score,0)=0 And zbid=@id";

                        MyData.MyCommand mcd = new MyData.MyCommand();
                        mcd.Execute(SQLDelete, mdpID);

                        mTableWorkspaceInspectCount++;
                    }
                }
                else
                {
                    MyRecord.Say("没有获取到任何内容");
                }
                MyRecord.Say("审核纪律单，完成");
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
            #endregion

        }

        #endregion



        #region 发送昨日还没处理的不良品退料和不合格品判定

        class SendIPQCAndScrapBackEmailItem
        {
            public string DepartmentName { get; set; }
            public int CountNumber { get; set; }
        }

        void SendIPQCAndScrapBackEmailLoder()
        {
            MyRecord.Say("发送昨日还没处理的不良品退料和不合格品判定..........");
            Thread t = new Thread(SendIPQCAndScrapBackEmail);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("发送昨日还没处理的不良品退料和不合格品判定。");
        }

        void SendIPQCAndScrapBackEmail()
        {
            try
            {
                MyRecord.Say("-----------------发送昨日还没处理的不良品退料和不合格品判定-------------------------");
                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{3}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=5 color=#ff0000 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容为截止到昨天（{0:yy/MM/dd HH}点前）各部门未判定和未退料的不良品单笔数列表。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; （说明：未判定笔数是指已经输入了完工单不良数，品保还没做不合格品处理单的完工单笔数。未退料笔数是指已经做了不合格品判定，还没做退料单的不合格品处理单笔数。）</FONT></DIV>
<DIV><FONT size=5 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
    <TR>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    部门    
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    未判定笔数
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    未退料笔数
    </TD>
    </TR>
    {1}
</TBODY></TABLE></FONT>
</DIV>
<DIV><FONT size=5 color=#ff33ff face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请品保和各部门尽快处理完成！</FONT></DIV>
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{2:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
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
    </TR>
";

                string SQLRS = @"Exec [_WH_IPQC_Return] -1";
                string SQLQC = @"Exec [_QC_IPQC_Confirm_View] 0,@BTime,@ETime,0,Null,Null";
                DateTime NowTime = DateTime.Now, BeginTime = DateTime.Parse("2016-01-01 00:00:05");
                MyData.MyParameter[] mps = new MyData.MyParameter[]
                {
                    new MyData.MyParameter("@BTime",BeginTime , MyData.MyParameter.MyDataType.DateTime),
                    new MyData.MyParameter("@ETime",NowTime , MyData.MyParameter.MyDataType.DateTime)
                };

                string brs = "";
                MyRecord.Say("后台计算");
                MyData.MyDataTable mrs = new MyData.MyDataTable(SQLRS, mps), mqc = new MyData.MyDataTable(SQLQC, mps);
                MyRecord.Say("后台计算完成。");
                try
                {
                    List<SendIPQCAndScrapBackEmailItem> lrs = new List<SendIPQCAndScrapBackEmailItem>();
                    List<SendIPQCAndScrapBackEmailItem> lqc = new List<SendIPQCAndScrapBackEmailItem>();
                    if (mrs != null && mrs.MyRows.Count > 0)
                    {
                        MyRecord.Say(string.Format("未退料，找到了：{0} 行。",mrs.MyRows.Count));
                        var v1 = from a in mrs.MyRows
                                 group a by Convert.ToString(a["department"]) into g
                                 select new SendIPQCAndScrapBackEmailItem
                                 {
                                     DepartmentName = g.Key,
                                     CountNumber = g.Count()
                                 };
                        lrs = v1.ToList();
                    }
                    if (mqc != null && mqc.MyRows.Count > 0)
                    {
                        MyRecord.Say(string.Format("未判定，找到了：{0} 行。", mqc.MyRows.Count));
                        var v2 = from a in mqc.MyRows
                                 group a by Convert.ToString(a["DepartmentName"]) into g
                                 select new SendIPQCAndScrapBackEmailItem
                                 {
                                     DepartmentName = g.Key,
                                     CountNumber = g.Count()
                                 };
                        lqc = v2.ToList();
                    }

                    var v3 = (from a in lqc
                              where (from b in lrs where b.DepartmentName == a.DepartmentName select b).Count() <= 0
                              select new { DeptName = a.DepartmentName, Count1 = a.CountNumber, Count2 = 0 }
                              ).Union(
                            (
                            from u in lrs
                            join m in lqc on u.DepartmentName equals m.DepartmentName
                            select new { DeptName = u.DepartmentName, Count1 = m.CountNumber, Count2 = u.CountNumber }
                            )
                            );
                    foreach (var item in v3)
                    {
                        brs += string.Format(br, item.DeptName, item.Count1, item.Count2);
                    }
                    MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", v3.Count()));
                    MyRecord.Say("创建SendMail。");
                    MyBase.SendMail sm = new MyBase.SendMail();
                    MyRecord.Say("加载邮件内容。");
                    sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, NowTime.AddDays(-1).Date.AddHours(20), brs, NowTime, MyBase.CompanyTitle));
                    sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_待处理不合格单和待处理退料单提醒。", NowTime, MyBase.CompanyTitle));
                    string mailto = ConfigurationManager.AppSettings["IPQCAndScrapBackMailTo"], mailcc = ConfigurationManager.AppSettings["IPQCAndScrapBackMailCC"];
                    MyRecord.Say(string.Format("\r\n MailTO:{0} \r\n MailCC:{1}", mailto, mailcc));
                    sm.MailTo = mailto; // "my18@my.imedia.com.tw,xang@my.imedia.com.tw,lghua@my.imedia.com.tw,my64@my.imedia.com.tw";
                    sm.MailCC = mailcc; // "jane123@my.imedia.com.tw,lwy@my.imedia.com.tw,my80@my.imedia.com.tw";
                    //sm.MailTo = "my80@my.imedia.com.tw";
                    MyRecord.Say("发送邮件。");
                    sm.SendOut();
                    MyRecord.Say("已经发送。");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                finally
                {

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
