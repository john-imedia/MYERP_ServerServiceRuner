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
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
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
    工单号    
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    料号
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    工序
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
    </TR>";

                string mailBodyMainSentence = string.Empty;
                string brs = string.Empty;
                vsMain = new C1FlexGrid();
                IssuesProduceList = null;
                IssuesProduceList = new List<FaceGridItem>();
                DateTime DBegin = DateTime.Now.AddDays(-1).Date.AddHours(23), DEnd = DateTime.Now.Date.AddDays(3).Date.AddHours(8);
                LoadGridSchame(DBegin, DEnd);
                SetError();
                int xSumIndex = SetGridTitleID();
                MyRecord.Say("创建SendMail。");
                MyBase.SendMail sm = new MyBase.SendMail();
                MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", vsMain.Rows.Count));
                string fname = string.Empty;
                if (xSumIndex > 0)
                {
                    mailBodyMainSentence = string.Format("发现工单交期在{1:yy/MM/dd HH:mm}-{2:yy/MM/dd HH:mm}，有{0}条工单排程有问题。详情请查看附档。", xSumIndex, DBegin, DEnd);
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
                    fname = string.Format("{0}\\{1}", dName, string.Format("NOTICE{0:yyyyMMdd}TMP{1}.xlsx", NowTime.Date, tmpFileNameLast));
                    MyRecord.Say(string.Format("fname = {0}", fname));
                    vsMain.AutoSizeCols();
                    vsMain.SaveExcel(fname, LCStr("出货计划和排程表"), FileFlags.NoFreezing | FileFlags.IncludeFixedCells | FileFlags.IncludeMergedRanges | FileFlags.VisibleOnly | FileFlags.SaveMergedRanges | FileFlags.OpenXml);
                    MyRecord.Say("已经保存了。");
                    sm.Attachments.Add(new System.Net.Mail.Attachment(fname));
                    vsMain.Dispose();
                    vsMain = null;
                    MyRecord.Say("加载到附件");
                    foreach (var item in IssuesProduceList)
                    {
                        brs += string.Format(br, item.ProduceRdsNo, item.ProductName, item.ProcessName, item.IssuesMemo);
                    }
                    brs = string.Format(gridTitle, brs);
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
                IssuesProduceList = null;
                MyRecord.Say("已经发送。");
                if (File.Exists(fname))
                {
                    File.Delete(fname);
                }
                MyRecord.Say("------------------发送完成----------------------------");
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        protected List<ProduceNoteItem> maindata, mainFinishData, MainPlanData;

        private C1FlexGrid vsMain;


        private class FaceGridItem
        {
            public string ProduceRdsNo { get; set; }
            public string ProductName { get; set; }

            public string ProcessName { get; set; }
            public string IssuesMemo { get; set; }
            public DateTime SendDate { get; set; }
            public DateTime PlanDate { get; set; }
        }

        private List<FaceGridItem> IssuesProduceList = new List<FaceGridItem>();

        private void LoadGridSchame(DateTime sDBegin, DateTime sDEnd)
        {

            try
            {
                #region 读取数据
                MyRecord.Say("开始计算。");
                MyRecord.Say(string.Format("开始时间：{0}，结束时间：{1}", sDBegin, sDEnd));
                MyRecord.Say("1，读取数据。");
                LoadDataSource(sDBegin, sDEnd);
                MyRecord.Say("2，处理数据。");
                MyRecord.Say("2.1，处理数据-v1");
                #endregion
                #region 处理数据
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
                            g.FirstOrDefault().StatusName
                        };
                MyRecord.Say("2.2，处理数据-v2");
                var v2 = from a in maindata
                         where a.SendDate >= sDBegin && a.SendDate <= sDEnd
                         group a by a.ProcessCode into g
                         orderby g.Key
                         select new
                         {
                             ProcessCode = g.Key,
                             ProcessName = g.FirstOrDefault().ProcessName
                         };
                #endregion
                MyRecord.Say("3，生成表格。");
                if (v.Count() <= 0)
                {
                    MyRecord.Say("3，v没有东西，退出。");
                    vsMain.Rows.Count = vsMain.Rows.Fixed;
                    return;
                }
                MyRecord.Say("3.1准备开始生产表格。");
                vsMain.Rows.Count = 1; vsMain.Cols.Count = 1;
                vsMain.Rows.Count = (v.Count() * 4) + 1; vsMain.Rows.Fixed = 1;
                vsMain.Cols.Count = v2.Count() + 11; vsMain.Cols.Fixed = 1; vsMain.Cols.Frozen = 8;
                vsMain.AllowMerging = AllowMergingEnum.RestrictAll; vsMain.AllowMergingFixed = AllowMergingEnum.RestrictAll;
                string[] titlewords = new string[] { "序号", "工单号", "产品号", "料号", "订单数", "交期", "需求数", "台号", "库存数", "有问题", "内容" };
                for (int i = 0; i < titlewords.Length; i++)
                {
                    vsMain[0, i] = LCStr(titlewords[i]);
                }
                vsMain.Cols[8].Name = "stocknumb";
                vsMain.Cols[8].Visible = false;
                vsMain.Cols[9].Name = "isError";
                vsMain.Cols[9].Visible = false;

                MyRecord.Say("3.2 开始加载表格数据。");
                int j = 1, u = 10;
                vsMain.Cols["isError"].DataType = typeof(int);

                CellStyle cstNoPlan = vsMain.Styles.Add("NoPlan", vsMain.Styles.Normal);
                cstNoPlan.BackColor = Color.FromArgb(255, 255, 0, 0);

                CellStyle cstStoped = vsMain.Styles.Add("Stoped", vsMain.Styles.Normal);
                cstStoped.BackColor = Color.FromArgb(255, 200, 200, 200);

                CellStyle cstFinished = vsMain.Styles.Add("Finished", vsMain.Styles.Normal);
                cstFinished.BackColor = Color.FromArgb(255, 196, 196, 0);

                CellStyle cstErrorPlan = vsMain.Styles.Add("ErrorPlan", vsMain.Styles.Normal);
                cstErrorPlan.BackColor = Color.FromArgb(255, 255, 0, 225);

                int bIndex = 1;
                foreach (var vitem in v)
                {
                    MyRecord.Say(string.Format("3.{0} 工单号： {1}", bIndex, vitem.ProduceNote));
                    for (int i = 0; i < 4; i++)
                    {
                        int jj = j + i;
                        vsMain[jj, 0] = vitem.ProduceNote;
                        vsMain[jj, 1] = string.Format("{0}\r\n({1})\r\n{2}", vitem.ProduceNote, vitem.StatusName, vitem.FinishRemark);
                        vsMain[jj, 2] = vitem.ProdCode;
                        vsMain[jj, 3] = vitem.ProdName;
                        vsMain[jj, 4] = vitem.ProductNumb;
                        vsMain[jj, 5] = string.Format("{0:MM/dd HH:ss}", vitem.SendDate);
                        vsMain[jj, 6] = vitem.pNumb;
                        vsMain[jj, 7] = vitem.PartID;
                        vsMain[jj, 8] = 0;
                        vsMain[jj, 9] = false;
                        vsMain.Rows[jj].UserData = vitem.ProduceNote;
                    }
                    u = 10;
                    vsMain[j, u] = LCStr("总完工数："); vsMain[j + 1, u] = LCStr("排期数：");
                    vsMain[j + 2, u] = LCStr("排期时间："); vsMain[j + 3, u] = LCStr("当日完工数：");

                    u++;
                    foreach (var v2Item in v2)
                    {
                        if (j == 1)
                        {
                            vsMain[0, u] = v2Item.ProcessName;
                        }

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
                                     BDD = a.SendDate
                                 };
                        var vpc = from a in MainPlanData
                                  where a.ProduceNote == vitem.ProduceNote && a.ProcessCode == v2Item.ProcessCode && a.PartID == vitem.PartID &&
                                        a.SendDate > DateTime.MinValue && a.SendDate <= vitem.SendDate
                                  select new
                                  {
                                      Numb1 = a.pNumb,
                                      BDD = a.SendDate
                                  };

                        if (vn.Count() > 0)
                        {
                            vsMain[j, u] = vn.Average(a => a.Numb1);
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
                                            vsMain[j + 1, u] = vp.OrderBy(a => a.BDD).FirstOrDefault().Numb1.ToString("f0");
                                            xDate = vp.Min(a => a.BDD);
                                            if (xDate > DateTime.Parse("2000-01-01"))
                                            {
                                                IssuesProduceList.Add(new FaceGridItem()
                                                {
                                                    ProduceRdsNo = vitem.ProduceNote,
                                                    ProductName = vitem.ProdName,
                                                    ProcessName = v2Item.ProcessName,
                                                    SendDate = vitem.SendDate,
                                                    PlanDate = xDate,
                                                    IssuesMemo = LCStr("排程时间有错误。")
                                                });
                                                for (int xi = 0; xi < 4; xi++)
                                                {
                                                    vsMain.SetCellStyle(j + xi, u, cstErrorPlan);
                                                    vsMain[j + xi, "isError"] = 1;
                                                }
                                                vsMain[j + 2, u] = string.Format("{0:MM/dd HH:mm}", xDate);
                                            }
                                            else
                                            {
                                                if (v2Item.ProcessCode != "9999")
                                                {
                                                    vsMain[j + 1, u] = LCStr("未排期");
                                                    vsMain[j + 2, u] = LCStr("未排期");
                                                    IssuesProduceList.Add(new FaceGridItem()
                                                        {
                                                            ProduceRdsNo = vitem.ProduceNote,
                                                            ProductName = vitem.ProdName,
                                                            ProcessName = v2Item.ProcessName,
                                                            SendDate = vitem.SendDate,
                                                            PlanDate = xDate,
                                                            IssuesMemo = LCStr("未排期")
                                                        });
                                                    
                                                    for (int xi = 0; xi < 4; xi++)
                                                    {
                                                        vsMain.SetCellStyle(j + xi, u, cstNoPlan);
                                                        vsMain[j + xi, "isError"] = 1;
                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            vsMain[j + 1, u] = vpc.OrderBy(a => a.BDD).FirstOrDefault().Numb1.ToString("f0");
                                            xDate = vpc.Min(a => a.BDD);
                                            vsMain[j + 2, u] = string.Format("{0:MM/dd HH:mm}", xDate);
                                        }
                                    }
                                    else
                                    {
                                        if (v2Item.ProcessCode != "9999")
                                        {
                                            vsMain[j + 1, u] = LCStr("未排期");
                                            vsMain[j + 2, u] = LCStr("未排期");
                                            IssuesProduceList.Add(new FaceGridItem()
                                            {
                                                ProduceRdsNo = vitem.ProduceNote,
                                                ProductName = vitem.ProdName,
                                                ProcessName = v2Item.ProcessName,
                                                SendDate = vitem.SendDate,
                                                PlanDate = xDate,
                                                IssuesMemo = LCStr("未排期")
                                            });
                                            for (int xi = 0; xi < 4; xi++)
                                            {
                                                vsMain.SetCellStyle(j + xi, u, cstNoPlan);
                                                vsMain[j + xi, "isError"] = 1;
                                            }
                                        }
                                    }
                                    var vf = from a in mainFinishData
                                             where a.ProduceNote == vitem.ProduceNote && a.ProcessCode == v2Item.ProcessCode && a.PartID == vitem.PartID &&
                                                   a.SendDate <= xDate
                                             select new
                                             {
                                                 Numb1 = a.FinishNumb,
                                                 Numb2 = a.RejectNumb + a.LossedNumb
                                             };
                                    if (vf.Count() > 0) vsMain[j + 3, u] = vf.Sum(a => a.Numb1);
                                }
                                else
                                {
                                    vsMain[j + 1, u] = LCStr("关闭");
                                    vsMain[j + 2, u] = vn.Max(a => a.Closer);
                                    vsMain[j + 3, u] = cv.ToString("MM-dd HH:mm");
                                    for (int xi = 0; xi < 4; xi++)
                                    {
                                        vsMain.SetCellStyle(j + xi, u, cstStoped);
                                    }
                                }
                            }
                            else
                            {
                                vsMain[j + 1, u] = LCStr("已完");
                                vsMain[j + 2, u] = LCStr("已完");
                                vsMain[j + 3, u] = LCStr("已完");
                                for (int xi = 0; xi < 4; xi++)
                                {
                                    vsMain.SetCellStyle(j + xi, u, cstFinished);
                                }
                            }
                        }
                        else
                        {
                            vsMain[j, u] = null;
                            vsMain[j + 1, u] = null;
                            vsMain[j + 2, u] = null;
                            vsMain[j + 3, u] = null;
                        }

                        u++;
                    }
                    j += 4;
                    bIndex++;
                }

                MyRecord.Say("4. 开始设置表格格式。");
                for (int i = 0; i < 12; i++) vsMain.Cols[i].AllowMerging = true;

                vsMain.Cols[8].Move(vsMain.Cols.Count - 1);
                vsMain.Rows[vsMain.Rows.Fixed - 1].AllowMerging = true;
                vsMain.Cols.Frozen = 7;
                vsMain.AutoSizeCols(vsMain.Cols.Frozen, vsMain.Cols.Count - 1, 1);
                vsMain.Cols[0].Width = 20;
                vsMain.Cols[1].Width = 80;
                vsMain.Cols[2].Width = vsMain.Cols[3].Width = 75;
                vsMain.Cols[4].Width = vsMain.Cols[6].Width = 55;
                vsMain.Cols[5].Width = 60;
                vsMain.Cols[7].Width = 35;
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        private int SetGridTitleID()
        {
            int u = 0, i = vsMain.Rows.Fixed;
            List<string> p = new List<string>();
            do
            {
                Row r = vsMain.Rows[i];
                string s = Convert.ToString(r.UserData);
                if (!p.Contains(s) && r.IsVisible)
                {
                    p.Add(s);
                    u += 1;
                }
                r.Caption = Convert.ToString(u);
                i++;
            } while (i < vsMain.Rows.Count);
            vsMain.AllowMergingFixed = AllowMergingEnum.RestrictAll;
            vsMain.Cols[0].AllowMerging = true;
            return u;
        }

        void SetError()
        {
            try
            {
                for (int j = 11; j < vsMain.Cols.Count; j++)
                {
                    vsMain.Cols[j].Visible = false;
                }
                for (int i = vsMain.Rows.Fixed; i < vsMain.Rows.Count; i++)
                {
                    bool v = (Convert.ToInt32(vsMain[i, "isError"]) == 1);
                    vsMain.Rows[i].Visible = v;
                    if (v)
                    {
                        for (int j = 11; j < vsMain.Cols.Count; j++)
                        {
                            bool cv = (vsMain[i, j] != null);
                            if ((!vsMain.Cols[j].Visible) && cv) vsMain.Cols[j].Visible = cv;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        protected virtual void LoadDataSource(DateTime DateBegin, DateTime DateEnd)
        {
            ///工单内容
            string SQL = @"
Set NoCount On

select a.ProdNo as ProdRdsNo,a.ProdID as ProductID,a.SendDate,a.pNumb,a.Inputer
Into #SendList
from _PMC_DeliverPlan_SendList a,moProduce b,pbProduct c
Where a.ProdId=b.id And b.Code=c.Code And (Not c.Type in ('115','116','1117')) 
      And a.SendDate Between @ProdBegin And @ProdEnd And b.InputDate < @ProdBegin
      And b.CloseDate is Null And b.CheckDate is Not Null And
      (c.Type=@ProdType or isNull(@ProdType,'')='') And
      (a.ProdNo= @ProductRdsNo or isNull(@ProductRdsNo,'')='') And
      (charindex(@ProdName,c.Name)>0 or isNull(@ProdName,'')='') And
      (b.CustID= @CustID or isNull(@CustID,'')='')

Select a.*,b.CustID as CustCode,b.Code as ProdCode,ProdName=Convert(nVarchar(100),Null),
       PartID=(Case When c.PartID='' Then '--' Else c.PartID End),
       c.ProcNo as ProcessCode,b.pNumb as ProductNumb,
       ProcessName=Convert(nVarchar(40),Null),
       c.FinishNumb,c.RejectNumb,c.LossedNumb,c.OverDate,
       SortID=Case When isNull(c.PartID,'')='' Then 110 else 100 End,
       SName=isNull((Select Name from moProdProperty Where Code=b.Property),'') + '，' +
             isNull((Select txt from ProdStatusView Where id=b.Status),'') + '，' +
             isNull((Select txt from ProdStockStatusView Where id=b.StockStatus),'未入库'),
       b.FinishRemark,c.CloseDate,c.Closer
Into #ProdList
from #SendList a,moProduce b,moProdProcedure c
Where a.ProductID=b.id And b.id=c.zbid And
      (c.ProcNo = @Process or isNull(@Process,'')='') And
      (c.Machinid in (select code from moMachine Where DepartmentID = @DeptID) or isNull(@DeptID,0)=0)

update a set a.ProcessName=x.Name from #ProdList a,moProcedure x Where a.ProcessCode=x.Code and SortID between 100 and 110

Insert Into #ProdList(ProdRdsNo,ProcessCode,ProcessName,FinishNumb,SendDate,SortID,PartID,CustCode,ProdCode,ProductNumb,SName,FinishRemark,pNumb)
Select  a.ProdRdsNo,'9999','成品入库',0,a.SendDate,500,'--' as PartID,b.CustID as CustCode,
        b.Code as ProdCode,b.pNumb as ProductNumb,       
        SName=isNull((Select Name from moProdProperty Where Code=b.Property),'') + '，' +
             isNull((Select txt from ProdStatusView Where id=b.Status),'') + '，' +
             isNull((Select txt from ProdStockStatusView Where id=b.StockStatus),'未入库'),
        b.FinishRemark,a.pNumb
From #SendList a,moProduce b Where a.ProductID=b.id And a.ProdRdsNo in (Select ProdRdsNo from #ProdList)

Update a set FinishNumb=(Select Sum(Numb) From stprdstocklst b Where b.ProductNo=a.ProdRdsNo) from #ProdList a where a.SortID=500
Update a Set a.ProdName=x.Name From #ProdList a,pbProduct x Where a.ProdCode=x.Code
Select *,b.Name as ProdName From #ProdList a Left Outer Join AllMaterialView b ON a.ProdCode=b.Code

Drop Table #SendList
Drop Table #ProdList
Set NoCount Off
";

            MyData.MyParameter[] mp = new MyData.MyParameter[]
            {
                new MyData.MyParameter("@ProdBegin",DateBegin, MyData.MyParameter.MyDataType.DateTime),
                new MyData.MyParameter("@ProdEnd",DateEnd, MyData.MyParameter.MyDataType.DateTime),
                new MyData.MyParameter("@Process",null),
                new MyData.MyParameter("@DeptID",0, MyData.MyParameter.MyDataType.Int),
                new MyData.MyParameter("@ProductRdsNo",null),
                new MyData.MyParameter("@CustID",null),
                new MyData.MyParameter("@ProdType",null),
                new MyData.MyParameter("@ProdName",null)
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
       ProductID=0,TolNumb as pNumb,
       FinishNumb = Numb1,RejectNumb=Numb2,LossedNumb=a.AdjustNumb + a.SampleNumb,
       ProcessCode=ProcessID,a.CustCode,SendDate=a.StartTime,ProdName=b.Name
from ProdDailyReport a Left Outer Join AllMaterialView b ON a.ProductCode=b.Code
Where a.ProduceNo in (Select ProdNo From _PMC_DeliverPlan_SendList) And a.EndTime Between @ProdBegin And @ProdEnd And
      (CharIndex('-' + @ProdType +'-',a.ProductCode)>0 or isNull(@ProdType,'')='') And
      (a.ProduceNo= @ProductRdsNo or isNull(@ProductRdsNo,'')='') And
      (a.CustCode= @CustID or isNull(@CustID,'')='') And
      (a.ProcessID = @Process or isNull(@Process,'')='') And
      (a.Machinid in (select code from moMachine Where DepartmentID = @DeptID) or isNull(@DeptID,0)=0)
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
Select b.PRODNO as ProdRdsNo,a.ProdCode,PartID=Case When a.PartNo='' Then '--' Else a.PartNo End,
       ProductID = b.PRODID,pNumb=a.PlanReqNumb,ProcessCode=a.ProcNo,a.CustCode,SendDate=a.Bdd,a.Closer,a.CloseDate,ProdName=c.Name
  from [_PMC_DeliverPlan_SendList] b Inner Join [_PMC_PlanProgressList_View] a ON a.zbid = b.PRODID
                                Left Outer Join AllMaterialView c ON a.ProdCode=c.Code
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

        protected class ProduceNoteItem
        {
            public ProduceNoteItem(MyData.MyDataRow mmdr)
            {
                ProduceNote = Convert.ToString(mmdr["ProdRdsNo"]);
                ProdCode = Convert.ToString(mmdr["ProdCode"]);
                ProdName = Convert.ToString(mmdr["ProdName"]);
                PartID = Convert.ToString(mmdr["PartID"]);
                ProductID = Convert.ToInt32(mmdr["ProductID"]);
                pNumb = Convert.ToInt32(mmdr["pNumb"]);
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
        }




        #endregion

    }
}
