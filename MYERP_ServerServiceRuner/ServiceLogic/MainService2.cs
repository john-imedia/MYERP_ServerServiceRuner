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

                        MyData.MyParameter mdpID=new MyData.MyParameter("@id",CurID, MyData.MyParameter.MyDataType.Int);
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
    }
}
