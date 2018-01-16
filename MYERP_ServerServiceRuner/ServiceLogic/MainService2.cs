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
            MyRecord.Say(string.Format("审核截止时间：{0:yy/MM/dd HH:mm}", NowTime));
            #region 审核纪律单
            try
            {
                MyRecord.Say("审核纪律单");
                SQL = @"Select * from _PMC_KanbanKPI_CheckList Where isNull(Status,0)=0 And InspectDate < @InputEnd";
                MyData.MyDataTable mTableWorkspaceInspect = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", NowTime, MyData.MyDataParameter.MyDataType.DateTime));
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

                        MyData.MyDataParameter mdpID = new MyData.MyDataParameter("@id", CurID, MyData.MyDataParameter.MyDataType.Int);
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
            public string FullSortID { get; set; }
            public int CountNumber { get; set; }
        }

        class SendIPQCAndScrapBackEmailItemRow
        {
            public string DepartmentName { get; set; }
            public int CountNumber1 { get; set; }
            public int CountNumber2 { get; set; }

            public int CountNumber3 { get; set; }

            public int CountNumber4 { get; set; }

            public int CountNumber5 { get; set; }
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
<DIV><FONT size=3 face=PMingLiU>{2}ERP系统提示您：</FONT></DIV>
{0}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{1:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");
                string gd = @"
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
    不良品品未判定笔数
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    不良品未退料笔数
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    不良品已退料未审核笔数
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    过版纸未退料笔数
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    过版纸已退料未审核笔数
    </TD>
    </TR>
    {1}
    </TBODY>
</TABLE>
</FONT>
</DIV>
<DIV><FONT size=5 color=#ff33ff face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请品保和各部门尽快处理完成！</FONT></DIV>
";
                string sd = @"
<DIV><FONT size=5 color=#ff0000 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 截止到昨天（{0:yy/MM/dd HH}点前）系统没有发现未处理的制程不良品和过版纸内容。</FONT></DIV>
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
    </TR>
";

                string SQLRS = @"Exec [_WH_IPQC_Return] -1";   //不良品退料单
                string SQLQC = @"Exec [_QC_IPQC_Confirm_View] 0,@BTime,@ETime,0,Null,Null,Null,Null,Null";   //不良品判定单
                string SQLRS1 = @"Select a.RdsNo,b.Name as Department,a.InputDate,a.Inputer,a.StockDate,a.StoreKeeper,a.QCChecker,a.Sender,a.Remark,a.inBonded,a.isBK,a.[Type],departmentFullSortID=b.FullSortID,
       BondedWord=(Case When a.inBonded =1 Then '保稅' else '非稅' end),
	   BKWord = (Case When a.isBK =1 Then 'BK' else '' end)
from [_WH_SemiReject_Receipt] a Inner Join [pbDept] b On a.DeptID = b.[_ID]
Where isNull(a.Type,'') = '' And a.CheckDate is Null And a.InputDate > '2016-10-01' And a.InputDate > @BTime"; //已退料未审核
                string SQLGB1 = @"Exec [_WH_IPQC_ReturnScrap_View2] 'N',@BTime,@ETime,0,Null,1,Null,Null,Null,-1,Null"; //未退过版纸
                string SQLGB2 = @"Select a.RdsNo,b.Name as Department,a.InputDate,a.Inputer,a.StockDate,a.StoreKeeper,a.QCChecker,a.Sender,a.Remark,a.inBonded,a.isBK,a.[Type],departmentFullSortID=b.FullSortID,
                                  BondedWord=(Case When a.inBonded =1 Then '保稅' else '非稅' end),
	                              BKWord = (Case When a.isBK =1 Then 'BK' else '' end)
                                  from [_WH_SemiReject_Receipt] a Inner Join [pbDept] b On a.DeptID = b.[_ID]
Where isNull(a.Type,'') = 'N' And a.CheckDate is Null And a.InputDate > '2016-10-01' And a.InputDate > @BTime"; //已退过版纸未审核
                DateTime NowTime = DateTime.Now, BeginTime = DateTime.Parse("2016-01-01 00:00:05");
                if (CompanyType == "MD") BeginTime = DateTime.Parse("2017-06-01 00:00:05");
                MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
                {
                    new MyData.MyDataParameter("@BTime",BeginTime , MyData.MyDataParameter.MyDataType.DateTime),
                    new MyData.MyDataParameter("@ETime",NowTime , MyData.MyDataParameter.MyDataType.DateTime)
                };

                string brs = "";
                MyRecord.Say("后台计算-退料");
                MyData.MyDataTable mrs = new MyData.MyDataTable(SQLRS, mps);   //退料
                MyRecord.Say("后台计算-不良判定");
                MyData.MyDataTable mqc = new MyData.MyDataTable(SQLQC, mps);   //判定
                MyRecord.Say("后台计算-退料未审核");
                MyData.MyDataTable mrs2 = new MyData.MyDataTable(SQLRS1, mps);   //退料未审核
                MyRecord.Say("后台计算-过版纸未退料");
                MyData.MyDataTable mgb1 = new MyData.MyDataTable(SQLGB1, mps);   //过版纸未退料
                MyRecord.Say("后台计算-过版纸退料未审核");
                MyData.MyDataTable mgb2 = new MyData.MyDataTable(SQLGB2, mps);   //过版纸退料未审

                MyRecord.Say("后台计算完成。");
                try
                {
                    List<SendIPQCAndScrapBackEmailItem> lrs = new List<SendIPQCAndScrapBackEmailItem>();
                    List<SendIPQCAndScrapBackEmailItem> lrs2 = new List<SendIPQCAndScrapBackEmailItem>();
                    List<SendIPQCAndScrapBackEmailItem> lqc = new List<SendIPQCAndScrapBackEmailItem>();
                    List<SendIPQCAndScrapBackEmailItem> lgb1 = new List<SendIPQCAndScrapBackEmailItem>();
                    List<SendIPQCAndScrapBackEmailItem> lgb2 = new List<SendIPQCAndScrapBackEmailItem>();
                    HSSFWorkbook wb = new HSSFWorkbook();
                    if (mrs != null && mrs.MyRows.Count > 0)  //退料
                    {
                        MyRecord.Say(string.Format("未退料，找到了：{0} 行。", mrs.MyRows.Count));
                        var v1 = from a in mrs.MyRows
                                 group a by a.Value("department") into g
                                 select new SendIPQCAndScrapBackEmailItem
                                 {
                                     DepartmentName = g.Key,
                                     FullSortID = g.FirstOrDefault().Value("DepartmentFullSortID"),
                                     CountNumber = g.Count()
                                 };
                        lrs = v1.ToList();
                        string[] fNames = new string[] { "department", "QCEditDate", "ProduceRdsNo", "FinishRdsNo", "RejectRdsNo", "ProdCode", "ProdName", "ProcessName", "MachineName", "ProjectName", "Number" };
                        string[] fTitles = new string[] { "部门", "判定日期", "工单号", "完工单号", "不良判定单号", "产品编号", "料号", "工序", "机台", "不良项目", "不良数" };
                        string[] fSortFields = new string[] { "departmentFullSortID", "ProcessCode", "ProduceRdsNo" };
                        ExportToExcelIPQCAndScrapBack(wb, mrs, fNames, fTitles, "未退料的制程不良品", fSortFields);
                    }
                    else
                    {
                        MyRecord.Say("没有未退料");
                    }
                    if (mqc != null && mqc.MyRows.Count > 0)  //不良品判定
                    {
                        MyRecord.Say(string.Format("未判定，找到了：{0} 行。", mqc.MyRows.Count));
                        var v2 = from a in mqc.MyRows
                                 group a by a.Value("DepartmentName") into g
                                 select new SendIPQCAndScrapBackEmailItem
                                 {
                                     DepartmentName = g.Key,
                                     FullSortID = g.FirstOrDefault().Value("DepartmentFullSortID"),
                                     CountNumber = g.Count()
                                 };
                        lqc = v2.ToList();

                        string[] fNames = new string[] { "DepartmentName", "RptDate", "ProduceNo", "AccNoteRdsNo", "ProdCode", "ProdName", "PartID", "ProcName", "MachineName", "RejNumb" };
                        string[] fTitles = new string[] { "部门", "日期", "工单号", "完工单号", "产品编号", "料号", "部件", "工序", "机台", "不良数" };
                        string[] fSortFields = new string[] { "DepartmentFullSortID", "ProcCode", "ProduceNo" };
                        ExportToExcelIPQCAndScrapBack(wb, mqc, fNames, fTitles, "制程不良品品保未判定的完工单", fSortFields);
                    }
                    else
                    {
                        MyRecord.Say("没有未判定");
                    }
                    if (mrs2 != null && mrs2.MyRows.Count > 0)  //已退料未审核
                    {
                        MyRecord.Say(string.Format("退料未审核，找到了：{0} 行。", mrs2.MyRows.Count));
                        var v1 = from a in mrs2.MyRows
                                 group a by a.Value("department") into g
                                 select new SendIPQCAndScrapBackEmailItem
                                 {
                                     DepartmentName = g.Key,
                                     FullSortID = g.FirstOrDefault().Value("DepartmentFullSortID"),
                                     CountNumber = g.Count()
                                 };
                        lrs2 = v1.ToList();
                        //				StockDate	StoreKeeper				inBonded	isBK	Type

                        string[] fNames = new string[] { "RdsNo", "Department", "InputDate", "Inputer", "QCChecker", "Sender", "Remark", "BondedWord", "BKWord" };
                        string[] fTitles = new string[] { "退料单号", "部门", "输入时间", "输入人", "品保", "退料人", "备注", "保税", "BK"};
                        string[] fSortFields = new string[] { "departmentFullSortID", "RdsNo", "Department" };
                        ExportToExcelIPQCAndScrapBack(wb, mrs2, fNames, fTitles, "退料单未审核的制程不良品", fSortFields);
                    }
                    else
                    {
                        MyRecord.Say("没有未审核");
                    }
                    if (mgb1 != null && mgb1.MyRows.Count > 0)  //退料
                    {
                        MyRecord.Say(string.Format("过版纸未退料，找到了：{0} 行。", mgb1.MyRows.Count));
                        var v1 = from a in mgb1.MyRows
                                 group a by a.Value("department") into g
                                 select new SendIPQCAndScrapBackEmailItem
                                 {
                                     DepartmentName = g.Key,
                                     FullSortID = g.FirstOrDefault().Value("DepartmentFullSortID"),
                                     CountNumber = g.Count()
                                 };
                        lgb1 = v1.ToList();
                        string[] fNames = new string[] { "department", "ProduceNo", "ProcessName", "MachineName", "ProductCode", "ProdName", "Bonded", "SecretWord", "ProjectName", "accNoterdsNo", "RejectrdsNo", "ReceiptrdsNo", "ReceiptStockDate", "rejNumb" };
                        string[] fTitles = new string[] { "部门", "工单号", "工序", "机台", "产品编号", "料号", "保税", "BK", "不良项目", "完工单号", "不合格品单", "退料单号", "退料时间", "退料数量" };
                        string[] fSortFields = new string[] { "departmentFullSortID", "ProcessID", "ProduceNo" };
                        ExportToExcelIPQCAndScrapBack(wb, mgb1, fNames, fTitles, "未退料的过版纸", fSortFields);
                    }
                    else
                    {
                        MyRecord.Say("没有过版纸未退料");
                    }

                    if (mgb2 != null && mgb2.MyRows.Count > 0)  //过版纸退料未审核
                    {
                        MyRecord.Say(string.Format("过版纸退料未审核，找到了：{0} 行。", mgb2.MyRows.Count));
                        var v1 = from a in mgb2.MyRows
                                 group a by a.Value("department") into g
                                 select new SendIPQCAndScrapBackEmailItem
                                 {
                                     DepartmentName = g.Key,
                                     FullSortID = g.FirstOrDefault().Value("DepartmentFullSortID"),
                                     CountNumber = g.Count()
                                 };
                        lgb2 = v1.ToList();
                        string[] fNames = new string[] { "RdsNo", "Department", "InputDate", "Inputer", "QCChecker", "Sender", "Remark", "BondedWord", "BKWord" };
                        string[] fTitles = new string[] { "退料单号", "部门", "输入时间", "输入人", "品保", "退料人", "备注", "保税", "BK" };
                        string[] fSortFields = new string[] { "departmentFullSortID", "RdsNo", "Department" };
                        ExportToExcelIPQCAndScrapBack(wb, mgb2, fNames, fTitles, "过版纸已退料未审核", fSortFields);
                    }
                    else
                    {
                        MyRecord.Say("过版纸没有退料未审核");
                    }



                    MyRecord.Say("生成邮件内容");
                    var vv = (from a in lqc
                              select a).Union(
                              from b in lrs
                              select b).Union(
                              from c in lrs2
                              select c).Union(
                              from d in lgb1
                              select d).Union(
                              from e in lgb2
                              select e);

                    var vvv = from a in vv
                              orderby a.FullSortID
                              group a by a.DepartmentName into g
                              select g.Key;

                    List<SendIPQCAndScrapBackEmailItemRow> xlist = new List<SendIPQCAndScrapBackEmailItemRow>();

                    foreach (var item in vvv)
                    {
                        SendIPQCAndScrapBackEmailItemRow xritem = new SendIPQCAndScrapBackEmailItemRow();
                        xritem.DepartmentName = item;
                        var vv1 = from a in lqc
                                  where a.DepartmentName == item
                                  select a.CountNumber;
                        if (vv1.IsNotEmptySet())
                        {
                            xritem.CountNumber1 = vv1.Sum();   //不合格品
                        }
                        var vv2 = from a in lrs
                                  where a.DepartmentName == item
                                  select a.CountNumber;
                        if (vv2.IsNotEmptySet())
                        {
                            xritem.CountNumber2 = vv2.Sum();  //退料数
                        }

                        var vv3 = from a in lrs2
                                  where a.DepartmentName == item
                                  select a.CountNumber;
                        if (vv3.IsNotEmptySet())
                        {
                            xritem.CountNumber3 = vv3.Sum(); //退料未审核
                        }

                        var vv4 = from a in lgb1
                                  where a.DepartmentName == item
                                  select a.CountNumber;
                        if (vv4.IsNotEmptySet())
                        {
                            xritem.CountNumber4 = vv4.Sum(); //过版纸
                        }

                        var vv5 = from a in lgb2
                                  where a.DepartmentName == item
                                  select a.CountNumber;
                        if (vv5.IsNotEmptySet())
                        {
                            xritem.CountNumber5 = vv5.Sum(); //过版纸未审核
                        }
                        xlist.Add(xritem);
                    }

                    string xBodyString = string.Empty;
                    if (xlist.Count() > 0)
                    {
                        foreach (var item in xlist)
                        {
                            brs += string.Format(br, item.DepartmentName, item.CountNumber1, item.CountNumber2, item.CountNumber3, item.CountNumber4, item.CountNumber5);
                        }
                        string gds = string.Format(gd, NowTime.AddDays(-1).Date.AddHours(20), brs);
                        xBodyString = string.Format(body, gds, NowTime, MyBase.CompanyTitle);
                        MyRecord.Say("邮件内容生成完毕。");
                    }
                    else
                    {
                        string sds = string.Format(sd, NowTime.AddDays(-1).Date.AddHours(20));
                        xBodyString = string.Format(body, sds, NowTime, MyBase.CompanyTitle);
                        MyRecord.Say("无内容");
                    }
                    MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", xlist.Count()));

                    MyRecord.Say("创建SendMail。");
                    MyBase.SendMail sm = new MyBase.SendMail();
                    MyRecord.Say("加载邮件内容。");

                    //要导出的
                    string tmpFileName = Path.GetTempFileName();
                    MyRecord.Say(string.Format("tmpFileName = {0}", tmpFileName));
                    Regex rgx = new System.Text.RegularExpressions.Regex(@"(?<=tmp)(.+)(?=\.tmp)");
                    string tmpFileNameLast = rgx.Match(tmpFileName).Value;
                    MyRecord.Say(string.Format("tmpFileNameLast = {0}", tmpFileNameLast));
                    string dName = string.Format("{0}\\TMPExcel", Application.StartupPath);
                    if (!Directory.Exists(dName)) Directory.CreateDirectory(dName);
                    string fname = string.Format("{0}\\{1}", dName, string.Format("NOTICE_{0:yyyyMMdd}_tmp{1}.xls", NowTime.Date, tmpFileNameLast));
                    MyRecord.Say(string.Format("fname = {0}", fname));
                    if (xlist.Count() > 0)
                    {
                        if (File.Exists(fname)) File.Delete(fname);
                        using (FileStream fs = new FileStream(fname, FileMode.Create, FileAccess.Write))
                        {
                            wb.Write(fs);
                            MyRecord.Say("已经保存了。");
                        }

                        if (File.Exists(fname))
                        {
                            sm.Attachments.Add(new System.Net.Mail.Attachment(fname));
                            MyRecord.Say("加载到附件");
                        }
                        else
                        {
                            MyRecord.Say("没找到附件");
                        }
                    }
                    wb.Close();
                    sm.MailBodyText = MyConvert.ZH_TW(xBodyString);
                    sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_待处理不合格单和待处理退料单提醒。", NowTime, MyBase.CompanyTitle));
                    string mailto = ConfigurationManager.AppSettings["IPQCAndScrapBackMailTo"], mailcc = ConfigurationManager.AppSettings["IPQCAndScrapBackMailCC"];
                    MyRecord.Say(string.Format("\r\n MailTO:{0} \r\n MailCC:{1}", mailto, mailcc));
                    sm.MailTo = mailto; // "my18@my.imedia.com.tw,xang@my.imedia.com.tw,lghua@my.imedia.com.tw,my64@my.imedia.com.tw";
                    sm.MailCC = mailcc; // "jane123@my.imedia.com.tw,lwy@my.imedia.com.tw,my80@my.imedia.com.tw";
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
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                finally
                {

                }
                MyRecord.Say("-----------------发送昨日还没处理的不良品退料和不合格品判定 完成-------------------------");
            }
            catch (Exception e)
            {
                MyRecord.Say(e);
            }
        }

        void ExportToExcelIPQCAndScrapBack(HSSFWorkbook wb, MyData.MyDataTable mds, string[] fileds, string[] titles, string caption, string[] sortFields)
        {
            if (mds.MyRows.IsNotEmptySet())
            {
                var vListGridRow = from a in mds.MyRows
                                   orderby a.Value(sortFields[0]), a.Value(sortFields[1]), a.Value(sortFields[2])
                                   select a;
                MyRecord.Say(string.Format("输出表格{0}", caption));
                HSSFSheet st = (HSSFSheet)wb.CreateSheet(caption);
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
                blodfont1.FontHeightInPoints = 18;
                blodfont1.Boldweight = (short)FontBoldWeight.Bold;

                HSSFCellStyle cds3 = (HSSFCellStyle)wb.CreateCellStyle();
                cds3.CloneStyleFrom(cellNomalStyle);
                cds3.SetFont(blodfont1);

                HSSFRow rNoteCaption1 = (HSSFRow)st.CreateRow(1);
                rNoteCaption1.HeightInPoints = 25;
                ICell xCellCaption1 = rNoteCaption1.CreateCell(0);
                xCellCaption1.SetCellValue(LCStr(caption));
                xCellCaption1.CellStyle = cds3;
                st.AddMergedRegion(new CellRangeAddress(1, 1, 0, fileds.Length - 1));

                //MyRecord.Say(string.Format("输出表格{0}，表格已经创建，建立表头。", caption));

                HSSFRow rNoteCaption2 = (HSSFRow)st.CreateRow(2);
                for (int i = 0; i < fileds.Length; i++)
                {
                    ICell xCellCaption = rNoteCaption2.CreateCell(i);
                    xCellCaption.SetCellValue(LCStr(titles[i]));
                    xCellCaption.CellStyle = cellNomalStyle;
                    //MyRecord.Say(string.Format("输出表格{0}，标题行，{1}列", caption, LCStr(titles[i])));
                }

                MyRecord.Say(string.Format("输出表格{0}，表格已经创建，建立表头完成。", caption));

                int uu = 3;

                foreach (var item in vListGridRow)
                {
                    if (item.IsNotNull())
                    {
                        HSSFRow rNoteCaption3 = (HSSFRow)st.CreateRow(uu);
                        for (int i = 0; i < fileds.Length; i++)
                        {
                            string fieldname = fileds[i];
                            //MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列", caption, uu, fieldname));
                            ICell xCellCaption = rNoteCaption3.CreateCell(i);
                            if (mds.Columns.Contains(fieldname))
                            {
                                //MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列，列存在。", caption, uu, fieldname));
                                DataColumn dc = mds.Columns[fieldname];
                                if (dc.IsNotNull())
                                {
                                    if (dc.DataType == typeof(DateTime))
                                    {
                                        //MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列，日期列。", caption, uu, fieldname));
                                        xCellCaption.SetCellValue(item.DateTimeValue(fieldname));
                                        xCellCaption.CellStyle = cellDateStyle;
                                        st.SetColumnWidth(i, 20 * 256);
                                    }
                                    else if (dc.DataType == typeof(double) || dc.DataType == typeof(Single) || dc.DataType == typeof(int) || dc.DataType == typeof(decimal) || dc.DataType == typeof(Int16))
                                    {
                                        //MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列，數字列。", caption, uu, fieldname));
                                        xCellCaption.SetCellValue(item.Value<double>(fieldname));
                                        xCellCaption.CellStyle = cellIntStyle;
                                        st.SetColumnWidth(i, 15 * 256);
                                    }
                                    else
                                    {
                                        //MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列，字符列。", caption, uu, fieldname));
                                        xCellCaption.SetCellValue(LCStr(item.Value(fieldname)));
                                        xCellCaption.CellStyle = cellNomalStyle;
                                        if (st.GetColumnWidth(i) < (item.Value(fieldname).Length + 4) * 256)
                                        {
                                            st.SetColumnWidth(i, (item.Value(fieldname).Length + 4) * 256);
                                        }
                                    }
                                    //MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列，完成。", caption, uu, fieldname));
                                }
                            }
                            else
                            {
                                MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列，没有这个字段。", caption, uu, fieldname));
                            }
                        }
                        uu++;
                    }
                }
            }
            MyRecord.Say(string.Format("输出表格{0}完成。", caption));
        }

        #endregion

        #region 记录所有的看板要件

        void KanbanRecorderLoader()
        {
            MyRecord.Say("开启生产看板后台计算..........");
            Thread t = new Thread(new ThreadStart(KanbanRecorder));
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启生产看板后台计算成功。");
        }

        void KanbanRecorder()
        {
            string SaveFinishNumberSQL = @"
Set NoCount On
Select isNull(a.Numb1,0) + isNull(a.Numb2,0) + isNull(a.AdjustNumb,0) + isNull(a.SampleNumb,0) as Numb,
       Convert(VarChar(10),(Case When DatePart(""HH"",a.StartTime) Between 0 And 7 Then DateAdd(DD,-1,a.StartTime) Else a.StartTime End),121) as [Date],
	   (Case When DatePart(HH,a.StartTime) Between 8 And 19 Then 1 Else 2 End) as PlanType,
	   a.MachinID as MachinCode,a.ProcessID
Into #T
  from ProdDailyReport a
  Where a.EndTime Between @DateBegin And @DateEnd

Select t.ProcessID,t.MachinCode,t.Date,t.PlanType,Sum(t.Numb) as Numb
  Into #P
  From #T t
Group by t.ProcessID,t.MachinCode,t.Date,t.PlanType

Update a Set DailyFinishNumb = t.Numb From [_PMC_ProdPlan_YieldRate] a,#P t
 Where a.PlanType = t.PlanType
   And Convert(VarChar(10),a.PlanBegin,121) = Convert(VarChar(10),t.Date,121)
   And a.MachineCode = t.MachinCode --And a.ProcessCode = t.ProcessID

Drop Table #T
Drop Table #P
";
            string SaveRejectBackTimeSQL = @"
Set NoCount On

Select Convert(VarChar(10),(Case When DatePart(""HH"",a.RptDate) Between 0 And 7 Then DateAdd(DD,-1,a.RptDate) Else a.RptDate End),121) as [Date],
	   (Case When DatePart(HH,a.RptDate) Between 8 And 19 Then 1 Else 2 End) as PlanType,
	   a.MachinID as MachineCode,a.ProcessID
Into #T
  from ProdDailyReport a
 Where a.RptDate Between @DateBegin And @DateEnd
   And isNull(a.Numb1,0) = 0 And isNull(Reject,0) = 1

Select Count(*) TimeNumb,Date,PlanType,MachineCode,ProcessID Into #P From #T Group by Date,PlanType,MachineCode,ProcessID

Update a Set RejectTimeNumb = t.TimeNumb
    From [_PMC_ProdPlan_YieldRate] a,#P t
   Where a.PlanType = t.PlanType And Convert(VarChar(10),a.PlanBegin,121) = Convert(VarChar(10),t.Date,121) and
         a.MachineCode = t.MachineCode --And a.ProcessCode = t.ProcessID

Drop Table #T
Drop Table #P
";
            string SaveInspectScoreSQL = @"
Set NoCount ON
Select a.InspectDate,a.ClassType,b.Score,b.DepartmentCode,b.MachineCode,b.ProcessCode,c.FullSortID,c.[_ID] as DepartmentID,ck.Result1Value as Peak,ck.LevelNumber
Into #T
  From [_PMC_KanbanKPI_CheckList] a Inner Join [_PMC_KanbanKPI_CheckList_List] b ON a.id=b.zbid
                                    Left Outer Join [_PMC_Machine_CheckItemList] ck ON b.ItemCode = ck.ItemCode
                                    Left Outer Join pbDept c ON b.DepartmentCode = c.Code
  Where isNull(b.Score,0) <> 0 And a.InspectDate Between @DateBegin And @DateEnd 

Select InspectDate,ClassType,DepartmentID,MachineCode,ProcessCode,(Case When Count(*) = 0 Then 0 Else Sum(Score) /(Count(*) / 10.00) End) as Score Into #X 
From #T Group by InspectDate,ClassType,DepartmentID,MachineCode,ProcessCode

Update a Set InspectScore =  t.Score / 100.00
    From [_PMC_ProdPlan_YieldRate] a,#X t
   Where a.PlanType = t.ClassType And Convert(VarChar(10),a.PlanBegin,121) = Convert(VarChar(10),t.InspectDate,121) and
         a.MachineCode = t.MachineCode --And a.ProcessCode = t.ProcessID

Drop Table #T
Drop Table #X
";
            DateTime NowDateTime = DateTime.Now;
            MyRecord.Say(string.Format("向前推：{0}天。", CacluateKanbanDaySpanTimes));
            DateTime xDateBegin = NowDateTime.AddDays(-CacluateKanbanDaySpanTimes).Date.AddHours(8), xDateEnd = NowDateTime.AddDays(1).Date.AddHours(9);

            MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
            {
                new MyData.MyDataParameter("@DateBegin",xDateBegin, MyData.MyDataParameter.MyDataType.DateTime),
                new MyData.MyDataParameter("@DateEnd",xDateEnd, MyData.MyDataParameter.MyDataType.DateTime)
            };

            DateTime dtBegin = DateTime.Now;

            MyData.MyCommand mc = new MyData.MyCommand();

            mc.Add(SaveFinishNumberSQL, "X1", mps);
            mc.Add(SaveInspectScoreSQL, "X2", mps);
            mc.Add(SaveRejectBackTimeSQL, "X3", mps);
            MyRecord.Say("提交计算");
            if (mc.Execute())
            {
                MyRecord.Say(string.Format("提交完成，计算三项共耗时：{0:0.0000}分钟", (DateTime.Now - dtBegin).TotalMinutes));
                ProudctionKanBan_Reject rej = new ProudctionKanBan_Reject(xDateBegin, xDateEnd);
                dtBegin = DateTime.Now;
                MyRecord.Say("计算不良数");
                if (rej.Run())
                {
                    MyRecord.Say(string.Format("提交完毕，计算和分配不良数，共耗时：{0:0.0000}分钟", (DateTime.Now - dtBegin).TotalMinutes));
                }
            }
        }

        class ProudctionKanBan_Reject
        {
            public ProudctionKanBan_Reject(DateTime dBegin, DateTime dEnd)
            {
                DateBegin = dBegin;
                DateEnd = dEnd;
            }
            public DateTime DateBegin { get; set; }

            public DateTime DateEnd { get; set; }

            void Say(string Word)
            {
                MyRecord.Say(Word);
            }

            public bool Run()
            {
                Say("开始计算，读取数据。");
                LoadData();

                MachineRejectData = new ProudctionKanBan_Reject_DataSource();
                Say("分配");
                int xci = 1;
                foreach (var item in RejectDataSource)
                {
                    var v = from a in FinishDataSource
                            where a.ProduceNo == item.ProduceNo && a.PrdID == item.AssignToPrdID && a.ProcessCode == item.AssignToProcessCode && a.DepartmentID == item.AssignToDepartmentID
                            select a;
                    MyRecord.Say(string.Format("分配{0}，{1}", item.ProduceNo, item.Date));
                    xci++;
                    if (v.Count() > 0)
                    {
                        if (v.Count() == 1)
                        {
                            ProudctionKanBan_Reject_DataItem xItem = new ProudctionKanBan_Reject_DataItem();
                            ProudctionKanBan_Finish_DataItem fItem = v.FirstOrDefault();
                            xItem.MachinCode = fItem.MachinCode;
                            xItem.PlanType = fItem.PlanType;
                            xItem.ProduceNo = item.ProduceNo;
                            xItem.AssignToDepartmentID = item.AssignToDepartmentID;
                            xItem.AssignToPrdID = item.AssignToPrdID;
                            xItem.AssignToProcessCode = item.AssignToProcessCode;
                            xItem.Date = item.Date;
                            xItem.Numb = item.Numb;
                            MachineRejectData.Add(xItem);
                        }
                        else
                        {
                            var vMachine = from a in v
                                           group a by a.MachinCode into g
                                           select g.Key;
                            double xmRejectNumb = item.Numb / vMachine.Count();
                            foreach (var mchItem in vMachine)
                            {
                                var vPlanType = from a in v
                                                where a.MachinCode == mchItem
                                                group a by a.PlanType into g
                                                select g.Key;
                                double xPTRejectNumb = xmRejectNumb / vPlanType.Count();
                                foreach (var PlanTypeItem in vPlanType)
                                {
                                    ProudctionKanBan_Reject_DataItem xItem = new ProudctionKanBan_Reject_DataItem();
                                    xItem.MachinCode = mchItem;
                                    xItem.PlanType = PlanTypeItem;
                                    xItem.ProduceNo = item.ProduceNo;
                                    xItem.AssignToDepartmentID = item.AssignToDepartmentID;
                                    xItem.AssignToPrdID = item.AssignToPrdID;
                                    xItem.AssignToProcessCode = item.AssignToProcessCode;
                                    xItem.Date = item.Date;
                                    xItem.Numb = xPTRejectNumb;
                                    MachineRejectData.Add(xItem);
                                }
                            }
                        }
                    }
                }
                Say("计算生成");
                var vMachineReject = from a in MachineRejectData
                                     group a by new { a.Date, a.PlanType, a.AssignToProcessCode, a.AssignToDepartmentID, a.MachinCode } into g
                                     select new
                                     {
                                         DepmartmentID = g.Key.AssignToDepartmentID,
                                         ProcessCode = g.Key.AssignToProcessCode,
                                         g.Key.MachinCode,
                                         g.Key.Date,
                                         g.Key.PlanType,
                                         RejectNumb = g.Sum(x => x.Numb)
                                     };

                MyData.MyCommand mcd = new MyData.MyCommand();
                Say("保存到表");
                string SQLDelete = @"Update [_PMC_ProdPlan_YieldRate] Set DailyRejectNumb = 0 Where DateDiff(""DD"",@BeginDate,PlanBegin) >= 0 ";
                mcd.Add(SQLDelete, "XDelete", new MyData.MyDataParameter("@BeginDate", DateBegin, MyData.MyDataParameter.MyDataType.DateTime));

                int vmiCount = 0;
                foreach (var item in vMachineReject)
                {
                    string SQLUpdate = @"
Declare @xid int
Select @xid=Max([_id]) From [_PMC_ProdPlan_YieldRate] Where DateDiff(""DD"",PlanBegin,@Date) >= 0 And DateDiff(""DD"",@BeginDate,PlanBegin) >=0 And DeaprtmentID = @DeptID And ProcessCode = @PCode And MachineCode = @MCode And PlanType = @PT
Update [_PMC_ProdPlan_YieldRate] Set DailyRejectNumb = isNull(DailyRejectNumb,0) + @Numb Where [_id] = @xid";
                    MyData.MyDataParameter[] mpUpdate = new MyData.MyDataParameter[]
                    {
                        new MyData.MyDataParameter("@MCode",item.MachinCode),
                        new MyData.MyDataParameter("@Date",item.Date, MyData.MyDataParameter.MyDataType.DateTime),
                        new MyData.MyDataParameter("@DeptID",item.DepmartmentID, MyData.MyDataParameter.MyDataType.Int),
                        new MyData.MyDataParameter("@PCode",item.ProcessCode),
                        new MyData.MyDataParameter("@PT",item.PlanType, MyData.MyDataParameter.MyDataType.Int),
                        new MyData.MyDataParameter("@Numb",item.RejectNumb, MyData.MyDataParameter.MyDataType.Numeric),
                        new MyData.MyDataParameter("@BeginDate",DateBegin, MyData.MyDataParameter.MyDataType.DateTime)
                    };
                    mcd.Add(SQLUpdate, string.Format("XM{0}", vmiCount), mpUpdate);
                }
                DateTime dtBegin = DateTime.Now;
                Say("开始提交");
                return mcd.Execute();
            }

            void LoadData()
            {
                MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
                {
                    new MyData.MyDataParameter("@DateBegin",DateBegin,MyData.MyDataParameter.MyDataType.DateTime),
                    new MyData.MyDataParameter("@DateEnd",DateEnd,MyData.MyDataParameter.MyDataType.DateTime)
                };

                string SQLIPQC = @"
Set NoCount On
Select p.ProduceNo,
       Convert(VarChar(10),(Case When DatePart(""HH"",a.EditDate) Between 0 And 7 Then DateAdd(DD,-1,a.EditDate) Else a.EditDate End),121) as [Date],
	   a.AssignToPrdID,a.AssignToProcessCode,a.AssignToDepartmentID,
	   a.RejNumb
Into #P
from [_PMC_IPQC_List] a Inner Join ProdDailyReport p ON a.zbid = p.[_ID]
Where a.EditDate Between @DateBegin And @DateEnd

Select a.Date,a.AssignToDepartmentID,a.AssignToPrdID,a.AssignToProcessCode,a.ProduceNo,Sum(a.RejNumb) as Numb
from #P a Group by a.Date,a.AssignToDepartmentID,a.AssignToPrdID,a.AssignToProcessCode,a.ProduceNo

Drop Table #P
";
                Say("读取IPQC判定明细。");
                RejectDataSource = new ProudctionKanBan_Reject_DataSource();
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQLIPQC, mps))
                {
                    if (md.MyRows.IsNotEmptySet())
                    {
                        var v = from a in md.MyRows
                                select new ProudctionKanBan_Reject_DataItem(a);
                        Say("已经读取，正在保存。");
                        RejectDataSource.AddRange(v.ToArray());
                    }
                }

                string SQLFinish = @"
Select isNull(a.Numb1,0) + isNull(a.Numb2,0) + isNull(a.AdjustNumb,0) + isNull(a.SampleNumb,0) as Numb,
	   (Case When DatePart(HH,a.StartTime) Between 8 And 19 Then 1 Else 2 End) as PlanType,
	   a.MachinID as MachinCode,a.ProcessID,a.PrdID,a.ProduceNo,m.DepartmentID as DepartmentID
Into #T
  from ProdDailyReport a Left Outer Join moMachine m On a.MachinID = m.code
  Where a.EndTime Between DateAdd(""MM"",-1,@DateBegin) And @DateEnd

Select MachinCode,ProcessID,PrdID,ProduceNo,DepartmentID,PlanType,Sum(Numb) as Numb
from #T Group by MachinCode,ProcessID,PrdID,ProduceNo,DepartmentID,PlanType
";
                Say("读取完工单明细。");
                FinishDataSource = new ProudctionKanBan_Finish_DataSource();
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQLFinish, mps))
                {
                    if (md.MyRows.IsNotEmptySet())
                    {
                        var v = from a in md.MyRows
                                select new ProudctionKanBan_Finish_DataItem(a);
                        Say("已经读取正在保存。");
                        FinishDataSource.AddRange(v.ToArray());
                    }
                }
            }

            public class ProudctionKanBan_Reject_DataItem
            {
                public ProudctionKanBan_Reject_DataItem(MyData.MyDataRow mr)
                {
                    if (mr.IsNotNull())
                    {
                        Date = mr.DateTimeValue("Date");
                        AssignToDepartmentID = mr.IntValue("AssignToDepartmentID");
                        AssignToPrdID = mr.IntValue("AssignToPrdID");
                        AssignToProcessCode = mr.Value("AssignToProcessCode");
                        ProduceNo = mr.Value("ProduceNo");
                        Numb = mr.Value<double>("Numb");
                    }
                }

                public ProudctionKanBan_Reject_DataItem()
                {

                }


                public DateTime Date { get; set; }

                public int AssignToDepartmentID { get; set; }

                public int AssignToPrdID { get; set; }

                public string AssignToProcessCode { get; set; }

                public string ProduceNo { get; set; }

                public double Numb { get; set; }

                public int PlanType { get; set; }

                public string MachinCode { get; set; }

            }

            public class ProudctionKanBan_Reject_DataSource : MyBase.MyEnumerable<ProudctionKanBan_Reject_DataItem>
            {

            }

            public ProudctionKanBan_Reject_DataSource RejectDataSource { get; set; }

            public ProudctionKanBan_Reject_DataSource MachineRejectData { get; set; }
            public class ProudctionKanBan_Finish_DataItem
            {
                public ProudctionKanBan_Finish_DataItem(MyData.MyDataRow r)
                {
                    MachinCode = r.Value("MachinCode");
                    ProcessCode = r.Value("ProcessID");
                    PrdID = r.IntValue("PrdID");
                    ProduceNo = r.Value("ProduceNo");
                    DepartmentID = r.IntValue("DepartmentID");
                    PlanType = r.IntValue("PlanType");
                    Numb = r.Value<double>("Numb");
                }

                public string MachinCode { get; set; }

                public string ProcessCode { get; set; }

                public int PrdID { get; set; }

                public string ProduceNo { get; set; }

                public int DepartmentID { get; set; }

                public int PlanType { get; set; }

                public double Numb { get; set; }
            }

            public class ProudctionKanBan_Finish_DataSource : MyBase.MyEnumerable<ProudctionKanBan_Finish_DataItem>
            {

            }

            public ProudctionKanBan_Finish_DataSource FinishDataSource { get; set; }
        }



        #endregion

        #region 清理LOG文件
        public void LogingFileCleanLoader()
        {
            Thread t = new Thread(new ThreadStart(LogingFileClean));
            t.IsBackground = true;
            t.Start();
        }
        public void LogingFileClean()
        {
            string LogsPath = string.Format("{0}\\Logs", Application.StartupPath);
            DirectoryInfo di = new DirectoryInfo(LogsPath);
            foreach (var item in di.GetFiles())
            {
                try
                {
                    if ((DateTime.Now - item.CreationTime).TotalDays > 1)
                    {
                        item.Delete();
                    }
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                    continue;
                }
            }
        }
        #endregion

    }
}
