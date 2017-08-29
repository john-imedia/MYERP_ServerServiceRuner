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
        void DeliverPlanNumbSumSender()
        {
            MyRecord.Say("开启发送每日出货计划量..........");
            Thread t = new Thread(DeliverPlanNumbSumSendMail);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("定时发送每日出货计划量已经启动。");
        }


        void DeliverPlanNumbSumSendMail()
        {
            try
            {
                MyRecord.Say("------------------开始定时发送7日出货计划量----------------------------");
                MyRecord.Say("从数据库搜寻内容");
                DateTime NowTime = DateTime.Now;
                string body = @"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#e6f3af #ffffff>
<DIV><FONT size=3 face=PMingLiU>{0}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;</FONT></DIV>
<DIV><FONT color=#ff0000 size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <B>{2:yy/MM/dd} {3}（详情请见附档）</B></FONT></DIV>
{4}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=2 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，切勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{1:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
";
                string gridTitle = @"
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
    <TR>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    日期
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    {0:yy/MM/dd}
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    {1:yy/MM/dd}
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    {2:yy/MM/dd}
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    {3:yy/MM/dd}
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    {4:yy/MM/dd}
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    {5:yy/MM/dd}
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    {6:yy/MM/dd}
    </TD>
    </TR>
    <TR>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    安排产量合计：
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
    </TR>
</TBODY></TABLE></FONT>
</DIV>";
                string mailBodyMainSentence = string.Empty;
                string brs = string.Empty;
                MyRecord.Say("读取数据");
                string SQL = @"
Declare @NowDate DateTime
Set @NowDate=Convert(DateTime,Convert(VarChar(10),GetDate(),121),121)
Select ProdNo = IsNull(A.ProdNo,B.RdsNo),SendDate = IsNull(A.SendDate,B.pdeliver),DateAdd([MI],-485,IsNull(A.SendDate,B.pDeliver)) as DutySendDate,PNumb=isNull(A.PNumb,B.PNumb),
       B.PNumb as ProduceNumber,B.Code,C.Name,C.Type,C.mTypeName as TypeName,C.CustCode
  Into #T
  from [_PMC_ActiveProduceNote_View] B Left Outer Join [_PMC_DeliverPlan_Sendlist] A ON B.[RdsNo] = A.PRODNO
                                       Left Outer Join [AllMaterialView] C ON B.Code = C.Code
Select * from #T a Where a.SendDate Between @NowDate And  DateAdd(""mi"",-1,DateAdd(""Day"",8,@NowDate))
Order by a.ProdNo
";
                List<DeilveryPlanItem> DataSource = new List<DeilveryPlanItem>();
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL))
                {
                    if (md.MyRows.IsNotEmptySet())
                    {
                        var v = from a in md.MyRows
                                select new DeilveryPlanItem(a);
                        DataSource = v.ToList();
                    }
                }
                MyRecord.Say("读取计算完成。");
                int xGridCount = DataSource.Count;
                MyRecord.Say("创建SendMail。");
                MyBase.SendMail sm = new MyBase.SendMail();
                string fname = string.Empty;
                if (xGridCount > 0)
                {
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
                    ExportToExcelDeliveryPlanNumber(fname, DataSource);
                    MyRecord.Say("已经保存了。");
                    sm.Attachments.Add(new System.Net.Mail.Attachment(fname));
                    MyRecord.Say("加载到附件");


                    var vSum = from a in DataSource
                               where a.DutySendDate.Date.Subtract(DateTime.Now.Date).TotalDays >= 0
                               group a by a.DutySendDate.Date into g
                               orderby g.Key ascending
                               select new
                               {
                                   Date = g.Key,
                                   Numb = string.Format("{0:#,##0.0}K", g.Sum(x => x.PNumb) / 1000.00)
                               };
                    DateTime[] dts = new DateTime[7]; string[] nbs = new string[7];
                    int dtsi = 0;
                    foreach (var item in vSum.Take(7))
                    {
                        dts[dtsi] = item.Date;
                        nbs[dtsi] = item.Numb;
                        dtsi++;
                    }
                    brs = string.Format(gridTitle, dts[0], dts[1], dts[2], dts[3], dts[4], dts[5], dts[6],
                                                   nbs[0], nbs[1], nbs[2], nbs[3], nbs[4], nbs[5], nbs[6]);
                    MyRecord.Say("生成邮件内容。");
                    mailBodyMainSentence = "未来7天出货计划安排。";
                }
                else
                {
                    mailBodyMainSentence = "没有出货计划内容。";
                    brs = string.Empty;
                }

                MyRecord.Say("加载邮件内容。");

                sm.MailBodyText = string.Format(body, MyBase.CompanyTitle, DateTime.Now, NowTime, mailBodyMainSentence, brs);
                sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}出货计划量", NowTime.Date, MyBase.CompanyTitle));
                string mailto = ConfigurationManager.AppSettings["DeliveryPlanNumberMailTo"], mailcc = ConfigurationManager.AppSettings["DeliveryPlanNumberMailCC"];
                MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mailto, mailcc));
                sm.MailTo = mailto;
                sm.MailCC = mailcc;
                //sm.MailTo = "my80@my.imedia.com.tw";
                MyRecord.Say("发送邮件。");
                sm.SendOut();
                sm.mail.Dispose();
                sm = null;
                MyRecord.Say("已经发送。");
                MyRecord.Say("------------------7日出货计划量-发送完成----------------------------");
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        class DeilveryPlanItem
        {
            public DeilveryPlanItem(MyData.MyDataRow mr)
            {
                ProdNo = mr.Value<string>("ProdNo");
                SendDate = mr.Value<DateTime>("SendDate");
                DutySendDate = mr.Value<DateTime>("DutySendDate");
                PNumb = mr.Value<double>("PNumb");
                ProduceNumber = mr.Value<double>("ProduceNumber");
                Code = mr.Value<string>("Code");
                Name = mr.Value<string>("Name");
                Type = mr.Value<string>("Type");
                TypeName = mr.Value<string>("TypeName");
            }

            public string ProdNo { get; set; }
            public DateTime SendDate { get; set; }
            public DateTime DutySendDate { get; set; }

            public double PNumb { get; set; }

            public double ProduceNumber { get; set; }

            public string Code { get; set; }

            public string Name { get; set; }

            public string Type { get; set; }

            public string TypeName { get; set; }
        }

        void ExportToExcelDeliveryPlanNumber(string FileName, List<DeilveryPlanItem> dataSource)
        {
            if (dataSource.IsNotEmptySet())
            {
                var v1 = from a in dataSource
                         where a.DutySendDate.Date.Subtract(DateTime.Now.Date).TotalDays >= 0
                         orderby a.Type
                         group a by a.TypeName into g
                         select g.Key;
                var v2 = from a in dataSource
                         where a.DutySendDate.Date.Subtract(DateTime.Now.Date).TotalDays >= 0
                         group a by a.DutySendDate.Date into g
                         orderby g.Key
                         select g.Key;

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
                cellDateStyle.DataFormat = wb.CreateDataFormat().GetFormat(@"yyyy/MM/dd");

                HSSFFont blodfont1 = (HSSFFont)wb.CreateFont();
                blodfont1.Boldweight = (short)FontBoldWeight.Bold;
                blodfont1.Color = HSSFColor.Turquoise.Index;

                HSSFFont blodfont2 = (HSSFFont)wb.CreateFont();
                blodfont2.Boldweight = (short)FontBoldWeight.Bold;
                blodfont2.Color = HSSFColor.DarkRed.Index;

                HSSFRow rGridCaption = (HSSFRow)st.CreateRow(0);
                ICell xCellCaption1 = rGridCaption.CreateCell(0);
                xCellCaption1.SetCellValue(LCStr("时间："));
                xCellCaption1.CellStyle = cellNomalStyle;
                int u = 0; int r = 0;
                foreach (var item in v1)
                {
                    r++;
                    HSSFRow rGridRow = (HSSFRow)st.CreateRow(r);
                    ICell xCellCaption = rGridRow.CreateCell(0);
                    xCellCaption.SetCellValue(item);
                    xCellCaption.CellStyle = cellNomalStyle;
                    u = 0;
                    foreach (var v2Item in v2)
                    {
                        var vv = from a in dataSource
                                 where a.DutySendDate.Date.Subtract(DateTime.Now.Date).TotalDays >= 0 &&
                                       a.TypeName == item && a.DutySendDate.Date == v2Item.Date
                                 select a.PNumb;
                        double vx = vv.Sum();
                        u++;
                        ICell xCell = rGridRow.CreateCell(u);
                        if (vx == 0)
                        {
                            xCell.SetCellValue("");
                        }
                        else
                        {
                            xCell.SetCellValue(string.Format("{0:#,##0.0}K", vx / 1000.00));
                        }
                        xCell.CellStyle = cellNomalStyle;
                    }
                    ICell xCellSum = rGridRow.CreateCell(u + 1);
                    var vvr = from a in dataSource
                              where a.DutySendDate.Date.Subtract(DateTime.Now.Date).TotalDays >= 0 && a.TypeName == item
                              select a.PNumb;
                    double vxr = vvr.Sum();
                    if (vxr == 0)
                    {
                        xCellSum.SetCellValue("");
                    }
                    else
                    {
                        xCellSum.SetCellValue(string.Format("{0:#,##0.0}K", vxr / 1000.00));
                    }
                    xCellSum.CellStyle = cellNomalStyle;
                }
                HSSFRow rSumRow = (HSSFRow)st.CreateRow(r + 1);
                u = 0;
                foreach (var v2Item in v2)
                {
                    u++;
                    ICell xCellCaption = rGridCaption.CreateCell(u);
                    xCellCaption.SetCellValue(v2Item);
                    xCellCaption.CellStyle = cellDateStyle;

                    var vv = from a in dataSource
                             where a.DutySendDate.Date.Subtract(DateTime.Now.Date).TotalDays >= 0 && a.DutySendDate.Date == v2Item.Date
                             select a.PNumb;
                    double vx = vv.Sum();

                    ICell xCelSum = rSumRow.CreateCell(u);
                    if (vx == 0)
                    {
                        xCelSum.SetCellValue("");
                    }
                    else
                    {
                        xCelSum.SetCellValue(string.Format("{0:#,##0.0}K", vx / 1000.00));
                    }
                    xCelSum.CellStyle = cellDateStyle;
                }

                ICell xCellCaption2 = rGridCaption.CreateCell(u + 1);
                xCellCaption2.SetCellValue(LCStr("合计"));
                xCellCaption2.CellStyle = cellDateStyle;

                ICell xCellCaption3 = rSumRow.CreateCell(0);
                xCellCaption3.SetCellValue(LCStr("合计"));
                xCellCaption3.CellStyle = cellDateStyle;

                var vvlast = from a in dataSource
                             where a.DutySendDate.Date.Subtract(DateTime.Now.Date).TotalDays >= 0
                             select a.PNumb;
                ICell xCellCaption4 = rSumRow.CreateCell(u + 1);
                double vxLast = vvlast.Sum();
                if (vxLast == 0)
                {
                    xCellCaption4.SetCellValue("");
                }
                else
                {
                    xCellCaption4.SetCellValue(string.Format("{0:#,##0.0}K", vxLast / 1000.00));
                }
                xCellCaption4.CellStyle = cellDateStyle;

                string filename = FileName;
                if (File.Exists(filename)) File.Delete(filename);
                using (FileStream fs = new FileStream(filename, FileMode.Create, FileAccess.Write))
                {
                    wb.Write(fs);
                }
            }

        }


        #endregion
    }
}
