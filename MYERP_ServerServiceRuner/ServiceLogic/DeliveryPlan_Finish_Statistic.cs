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
                           orderby a.DepartmentName
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
                        brs += string.Format(br, item.DepartmentName, item.ProcessName, item.ProductName, item.ProduceRdsNo, item.IssuesMemo);
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
                string mailto = ConfigurationManager.AppSettings["DeliveryPlanMailTo"], mailcc = ConfigurationManager.AppSettings["DeliveryPlanMailCC"];
                MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mailto, mailcc));
                sm.MailTo = mailto;
                sm.MailCC = mailcc;
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


        protected List<ProduceNoteItem> maindata, mainFinishData, MainPlanData;

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
Set NoCount On

select a.ProdNo as ProdRdsNo,a.ProdID as ProductID,a.SendDate,a.pNumb,a.Inputer
Into #SendList
from _PMC_DeliverPlan_SendList a Inner Join moProduce b On a.ProdId=b.id
                                 Inner Join pbProduct c On b.Code=c.Code
Where isNull(b.Status,0) >0 And b.CloseDate is Null And b.Finalized is Null And b.StockDate is Null And b.FinishDate is Null And
      (c.Type Not in ({0})) And a.SendDate Between @ProdBegin And @ProdEnd And b.InputDate < @ProdBegin And
      (c.Type=@ProdType or isNull(@ProdType,'')='') And
      (a.ProdNo= @ProductRdsNo or isNull(@ProductRdsNo,'')='') And
      (charindex(@ProdName,c.Name)>0 or isNull(@ProdName,'')='') And
      (b.CustID= @CustID or isNull(@CustID,'')='')
Create Index _Ix_SendList_ProdRdsNo on #SendList (ProdRdsNo)
Select a.*,b.CustID as CustCode,b.Code as ProdCode,ProdName=Convert(nVarchar(100),Null),
       PartID=(Case When c.PartID='' Then '--' Else c.PartID End),
       c.ProcNo as ProcessCode,b.pNumb as ProductNumb,
       ProcessName=Convert(nVarchar(40),Null),
       c.FinishNumb,c.RejectNumb,c.LossedNumb,c.OverDate,
       SortID=Case When isNull(c.PartID,'')='' Then 110 else 100 End,
       SName=isNull((Select Name from moProdProperty Where Code=b.Property),'') + '，' +
             isNull((Select txt from ProdStatusView Where id=b.Status),'') + '，' +
             isNull((Select txt from ProdStockStatusView Where id=b.StockStatus),'未入库'),
       b.FinishRemark,c.CloseDate,c.Closer,m.Code as MachineCode,m.name as MachineName,
	   DepartmentName = (Select Name from pbDept Where [_id] = m.DepartmentID)
Into #ProdList
from #SendList a Inner Join moProduce b ON a.ProductID=b.id
                 Inner Join moProdProcedure c ON b.id=c.zbid
				 Left Outer Join moMachine m On c.MachinID = m.Code
Where isNull(b.Status,0) >0 And b.CloseDate is Null And b.Finalized is Null And b.StockDate is Null And b.FinishDate is Null And
      (c.ProcNo = @Process or isNull(@Process,'')='') And
      (m.DepartmentID = @DeptID or isNull(@DeptID,0)=0)

Create Index _Ix_ProdList_ProdRdsNo on #ProdList (ProdRdsNo,SortID)
update a set a.ProcessName=x.Name from #ProdList a,moProcedure x Where a.ProcessCode=x.Code and SortID between 100 and 110

Insert Into #ProdList(ProdRdsNo,ProcessCode,ProcessName,FinishNumb,SendDate,SortID,PartID,CustCode,ProdCode,ProductNumb,SName,FinishRemark,pNumb)
Select  a.ProdRdsNo,'9999','成品入库',0,a.SendDate,500,'--' as PartID,b.CustID as CustCode,
        b.Code as ProdCode,b.pNumb as ProductNumb,
        SName=isNull((Select Name from moProdProperty Where Code=b.Property),'') + '，' +
             isNull((Select txt from ProdStatusView Where id=b.Status),'') + '，' +
             isNull((Select txt from ProdStockStatusView Where id=b.StockStatus),'未入库'),
        b.FinishRemark,a.pNumb
From #SendList a,moProduce b Where a.ProductID=b.id And a.ProdRdsNo in (Select ProdRdsNo from #ProdList)

Update a set FinishNumb=T.Numb from #ProdList a,(Select Sum(Numb) as Numb,b.ProductNo From stprdstocklst b Group by b.ProductNo) T where T.ProductNo=a.ProdRdsNo and a.SortID=500
Update a Set a.ProdName=x.Name From #ProdList a,pbProduct x Where a.ProdCode=x.Code
Select *,b.Name as ProdName,c.StockNumb From #ProdList a Left Outer Join AllMaterialView b ON a.ProdCode=b.Code
                                             Left Outer Join [_ST_StockListByCodeView] c On a.ProdCode = c.Code

Drop Table #SendList
Drop Table #ProdList
Set NoCount Off
";
            SQL = string.Format(SQL, ConfigurationManager.AppSettings["DeliveryPlanFinishStaticExceptProdTypes"]);
            MyData.MyDataParameter[] mp = new MyData.MyDataParameter[]
            {
                new MyData.MyDataParameter("@ProdBegin",DateBegin, MyData.MyDataParameter.MyDataType.DateTime),
                new MyData.MyDataParameter("@ProdEnd",DateEnd, MyData.MyDataParameter.MyDataType.DateTime),
                new MyData.MyDataParameter("@Process",null),
                new MyData.MyDataParameter("@DeptID",0, MyData.MyDataParameter.MyDataType.Int),
                new MyData.MyDataParameter("@ProductRdsNo",null),
                new MyData.MyDataParameter("@CustID",null),
                new MyData.MyDataParameter("@ProdType",null),
                new MyData.MyDataParameter("@ProdName",null)
            };
            DateTime StartTime = DateTime.Now;
            MyRecord.Say("1.0 准备开始...");
            using (MyData.MyDataTable mdata = new MyData.MyDataTable(SQL, mp))
            {
                var v = from a in mdata.MyRows
                        select new ProduceNoteItem(a);
                maindata = v.ToList();
            }
            MyRecord.Say(string.Format("1.1 读取工单，耗时：{0}秒。", (DateTime.Now - StartTime).TotalSeconds));
            StartTime = DateTime.Now;
            ///完工内容

            SQL = @"
Select ProduceNo as ProdRdsNo,ProdCode=ProductCode,PartID=Case When PartID='' Then '--' Else PartID End,
       ProductID=0,TolNumb as pNumb,m.Name as MachineName,DepartmentName=(Select Name From pbDept p Where p.[_ID]=m.DepartmentID),
       FinishNumb = Numb1,RejectNumb=Numb2,LossedNumb=AdjustNumb + SampleNumb,
       ProcessCode=ProcessID,CustCode,SendDate=a.StartTime
from  ProdDailyReport a Inner Join moMachine m On a.MachinID = m.Code
Where a.ProduceNo in (Select ProdNo From _PMC_DeliverPlan_SendList) And a.EndTime Between @ProdBegin And @ProdEnd And
      (CharIndex('-' + @ProdType +'-',ProductCode)>0 or isNull(@ProdType,'')='') And
      (ProduceNo= @ProductRdsNo or isNull(@ProductRdsNo,'')='') And
      (CustCode= @CustID or isNull(@CustID,'')='') And
      (ProcessID = @Process or isNull(@Process,'')='') And
      (m.DepartmentID = @DeptID or isNull(@DeptID,0)=0)
";
            using (MyData.MyDataTable mdata = new MyData.MyDataTable(SQL, mp))
            {
                var v = from a in mdata.MyRows
                        select new ProduceNoteItem(a);
                mainFinishData = v.ToList();
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
Select b.PRODNO as ProdRdsNo,a.ProdCode,PartID=Case When a.PartNo='' Then '--' Else a.PartNo End,a.DepartmentName,a.MachineName,
       ProductID = b.PRODID,pNumb=a.PlanReqNumb,ProcessCode=a.ProcNo,a.CustCode,SendDate=a.Bdd,a.Closer,a.CloseDate,ProdName=c.Name
  from [_PMC_DeliverPlan_SendList] b Inner Join [_PMC_PlanProgressList_View] a ON a.zbid = b.PRODID
                                     Inner Join AllMaterialView c On a.ProdCode = c.Code
 Where (a.ProcNo = @Process or isNull(@Process,'')='') And
       (a.DepartmentID =@DeptID or isNull(@DeptID,0)=0) And
       (CharIndex('-' + @ProdType +'-',a.ProdCode)>0 or isNull(@ProdType,'')='') And
       (b.ProdNo= @ProductRdsNo or isNull(@ProductRdsNo,'')='') And
       (a.CustCode= @CustID or isNull(@CustID,'')='')
";

            using (MyData.MyDataTable mdata = new MyData.MyDataTable(SQL, mp))
            {
                var v = from a in mdata.MyRows
                        select new ProduceNoteItem(a);
                MainPlanData = v.ToList();
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
                var v = from a in maindata
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
                var v2 = from a in maindata
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
                foreach (var vitem in v)
                {
                    uIndex++;
                    MyRecord.Say(string.Format("计算第{0}条，共{1}条，{2}。", uIndex, v.Count(), vitem.ProduceNote));
                    GridProduceNoteItem GridProdItem = new GridProduceNoteItem();
                    GridList.Add(GridProdItem);
                    GridProdItem.ID = uIndex;
                    GridProdItem.ProduceRdsNo = vitem.ProduceNote;
                    GridProdItem.ProduceNote = string.Format("{0}\r\n({1})\r\n{2}", vitem.ProduceNote, vitem.StatusName, vitem.FinishRemark);
                    GridProdItem.ProductCode = vitem.ProdCode;
                    GridProdItem.ProductName = vitem.ProdName;
                    GridProdItem.ProductNumb = vitem.ProductNumb;
                    GridProdItem.SendDate = vitem.SendDate;
                    GridProdItem.pNumb = vitem.pNumb;
                    GridProdItem.PartID = vitem.PartID;
                    GridProdItem.StockNumb = vitem.StockNumb;
                    GridProdItem.GridProcess = new List<GridProcessItem>();
                    foreach (var v2Item in v2)
                    {
                        var vn = from a in maindata
                                 where a.ProduceNote == vitem.ProduceNote && a.ProcessCode == v2Item.ProcessCode && a.PartID == vitem.PartID
                                 select new
                                 {
                                     Numb1 = a.FinishNumb,
                                     Numb2 = a.RejectNumb + a.LossedNumb,
                                     Overdate = a.OverDate,
                                     CloseDate = a.CloseDate,
                                     Closer = a.Closer
                                 };
                        var vp = from a in MainPlanData
                                 where a.ProduceNote == vitem.ProduceNote && a.ProcessCode == v2Item.ProcessCode && a.PartID == vitem.PartID
                                 select new
                                 {
                                     Numb1 = a.pNumb,
                                     BDD = a.SendDate,
                                     MachineName = a.MachineName,
                                     DepartmentName = a.DepartmentName
                                 };
                        var vpc = from a in MainPlanData
                                  where a.ProduceNote == vitem.ProduceNote && a.ProcessCode == v2Item.ProcessCode && a.PartID == vitem.PartID &&
                                        a.SendDate > DateTime.MinValue && a.SendDate <= vitem.SendDate
                                  select new
                                  {
                                      Numb1 = a.pNumb,
                                      BDD = a.SendDate,
                                      MachineName = a.MachineName,
                                      DepartmentName = a.DepartmentName
                                  };
                        if (vn.Count() > 0)
                        {
                            GridProcessItem gpi = new GridProcessItem();
                            GridProdItem.GridProcess.Add(gpi);
                            gpi.TotalFinishNumb = vn.Average(a => a.Numb1);
                            gpi.ProcessCode = v2Item.ProcessCode;
                            gpi.ProcessName = v2Item.ProcessName;
                            gpi.ProduceRdsNo = GridProdItem.ProduceRdsNo;
                            gpi.ProductName = GridProdItem.ProductName;

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
                                            xDate = vp.Min(a => a.BDD);
                                            if (xDate > DateTime.Parse("2000-01-01"))
                                            {
                                                gpi.GridStyle = GridProdItem.ErrorRowStyle = "ErrorPlan";
                                                gpi.isError = GridProdItem.isError = true;
                                                gpi.PlanDate = LCStr(string.Format("{0:MM/dd HH:mm}(错误)", xDate));
                                                xMachineName = vp.FirstOrDefault().MachineName;
                                                if (rgxMachineName.IsMatch(xMachineName)) xMachineName = rgxMachineName.Match(xMachineName).Value;
                                                gpi.PlanMachine = string.Format("{0}({1})", xMachineName, vp.FirstOrDefault().DepartmentName);
                                                gpi.DepartmentName = vp.FirstOrDefault().DepartmentName;
                                                gpi.IssuesMemo = LCStr("排程时间错误");
                                            }
                                            else
                                            {
                                                if (v2Item.ProcessCode != "9999")
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
                                        if (v2Item.ProcessCode != "9999")
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
            if (maindata.IsNotEmptySet() && GridList.IsNotEmptySet())
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

        protected class ProduceNoteItem
        {
            public ProduceNoteItem(MyData.MyDataRow mmdr)
            {
                ProduceNote = mmdr.Value("ProdRdsNo");
                ProdCode = mmdr.Value("ProdCode");
                ProdName = mmdr.Value("ProdName");
                PartID = mmdr.Value("PartID");
                ProductID = mmdr.IntValue("ProductID");
                pNumb = mmdr.IntValue("pNumb");
                FinishNumb = Convert.ToInt32(mmdr["FinishNumb"]);
                RejectNumb = Convert.ToInt32(mmdr["RejectNumb"]);
                LossedNumb = Convert.ToInt32(mmdr["LossedNumb"]);
                SendDate = Convert.ToDateTime(mmdr["SendDate"]);
                Inputer = Convert.ToString(mmdr["Inputer"]);
                CustCode = Convert.ToString(mmdr["CustCode"]);
                ProcessCode = Convert.ToString(mmdr["ProcessCode"]);
                ProcessName = Convert.ToString(mmdr["ProcessName"]);
                ProductNumb = Convert.ToInt32(mmdr["ProductNumb"]);
                SortID = Convert.ToInt32(mmdr["SortID"]);
                OverDate = Convert.ToDateTime(mmdr["OverDate"]);
                StatusName = Convert.ToString(mmdr["SName"]);
                FinishRemark = Convert.ToString(mmdr["FinishRemark"]);
                Closer = Convert.ToString(mmdr["Closer"]);
                CloseDate = Convert.ToDateTime(mmdr["CloseDate"]);
                DepartmentName = Convert.ToString(mmdr["DepartmentName"]);
                MachineName = Convert.ToString(mmdr["MachineName"]);
                StockNumb = mmdr.Value<double>("StockNumb");
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

            public bool isError { get; set; }

            public string IssuesMemo { get; set; }
            public string ProduceRdsNo { get; set; }
            public string ProductName { get; set; }
        }


        #endregion

    }
}
