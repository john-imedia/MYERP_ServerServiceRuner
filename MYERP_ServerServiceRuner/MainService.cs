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
using System.Windows.Forms;
using System.IO;
using MYERP_ServerServiceRuner.Base;
using NPOI;
using NPOI.Util;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace MYERP_ServerServiceRuner
{
    public partial class MainService : ServiceBase
    {
        private System.Timers.Timer MainTimer;
        #region 系统处理
        public MainService()
        {
            InitializeComponent();
            this.MainTimer = new System.Timers.Timer();
        }

        public DateTime CheckStockStartTime, CalculatePruchaseStartTime, CalculateOLDPayToNew, CalculateOrderStartTime, CheckPlanConfirmStartTime, KickOutSPUserStartTime, SendIPQCListStaticStartTime;
        /// <summary>
        /// 上一次计算完工数领料数计算时间。
        /// </summary>
        public DateTime ProduceFeedBackLastRunTime = DateTime.MinValue, ProduceEstimateLastRunTime = DateTime.MinValue;

        public static string CompanyType { get; set; }

        public double CheckStockTimers = 0;   //3.25;
        /// <summary>
        /// 计算达成率，向前计算多少天的设置。
        /// </summary>
        public int CacluatePlanRateDaySpanTimes = 6,
                   CacluateKanbanDaySpanTimes = 6,
                   CacluateOEEDaySpanTimes = 6;

        public bool FirstStart = false,
                    CheckStockNoteOnceTimeSwitch = false,
                    ProdKanbanSaveOnceTimeSwitch = false;

        protected override void OnStart(string[] args)
        {
            MainTimer.Interval = 1000;
            MainTimer.Elapsed += MainTimer_Elapsed;
            MyRecord.Say(MyBase.ConnectionString);
            MyRecord.Say(MyBase.CompanyTitle);
            CheckStockStartTime = DateTime.Now.Date;
            ProduceFeedBackLastRunTime = CalculatePruchaseStartTime = CalculateOrderStartTime = CheckPlanConfirmStartTime = KickOutSPUserStartTime = SendIPQCListStaticStartTime = DateTime.Now;
            //ProduceFeedBackLastRunTime = new DateTime(2000, 1, 1);
            MyRecord.Say("服务启动");
            FirstStart = true;
            MainTimer.Start();
        }
        #endregion

        void MainTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                CompanyType = MyConfig.ApplicationConfig.CompanyType;
                CheckStockTimers = MyConfig.ApplicationConfig.CheckStockTimers;
                CacluatePlanRateDaySpanTimes = MyConfig.ApplicationConfig.CacluatePlanRateDaySpanTimes;
                CacluateOEEDaySpanTimes = MyConfig.ApplicationConfig.CacluateOEEDaySpanTimes;
                CacluateKanbanDaySpanTimes = MyConfig.ApplicationConfig.CacluateKanbanDaySpanTimes;

                DateTime NowTime = DateTime.Now;
                int h = NowTime.Hour, m = NowTime.Minute, s = NowTime.Second, d = NowTime.Day;
                DayOfWeek w = NowTime.DayOfWeek;


                if (FirstStart)
                {
                    FirstStart = false;
                    MyRecord.Say("首次启动，时间循环已经开启。");
                    CheckStockStartTime = ProduceEstimateLastRunTime = NowTime;
                    ProduceFeedBackLastRunTime = NowTime.AddDays(-5);
                    MyRecord.Say(string.Format("设定时间成功：CheckStockStartTime={0} ProduceFeedBackLastRunTime = {1} ", CheckStockStartTime, ProduceFeedBackLastRunTime));
                    #region 暂停

                    #endregion
                }
                //审核库存单据
                CheckStockNoteOnceTimeSwitch = MyConfig.ApplicationConfig.CheckStockNoteOnceTime;
                if (!CheckStockRecordRunning && CheckStockNoteOnceTimeSwitch)
                {
                    CheckStockStartTime = NowTime;
                    MyRecord.Say("临时启动审核库存单据，已经启动");
                    CheckStockRecordRunning = true;
                    CheckStockRecordLoder();   //審核出入庫單。
                    Thread.Sleep(500);
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["CheckStockNoteOnceTime"].BoolValue = false;
                    cfg.SaveToFile(fileName);
                }
                CheckStockNoteOnceTimeSwitch = false;

                ProdKanbanSaveOnceTimeSwitch = MyConfig.ApplicationConfig.ProdKanbanSaveOnceTime;

                //临时计算月达成率和看板。计算限度根据设定。
                if (ProdKanbanSaveOnceTimeSwitch)
                {
                    ProdKanbanSaveOnceLoader(NowTime);
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["ProdKanbanSaveOnceTime"].BoolValue = false;
                    cfg.SaveToFile(fileName);
                }
                ProdKanbanSaveOnceTimeSwitch = false;

                //批量计算完工和领料数
                if (MyConfig.ApplicationConfig.CacluateProduceFeedBackTimes > 0)
                {
                    mCacluateProduceFeedBackTimes = MyConfig.ApplicationConfig.CacluateProduceFeedBackTimes;
                    ProduceFeedBackManaulLoder();
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["CacluateProduceFeedBackTimes"].IntValue = 0;
                    cfg.SaveToFile(fileName);
                }

                //发送差异数
                if (MyConfig.ApplicationConfig.SendProduceDiffNumbEmailOnce)
                {
                    SendProduceDiffNumbEmailLoder2();
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["SendProduceDiffNumbEmailOnce"].BoolValue = false;
                    cfg.SaveToFile(fileName);
                }

                //3日出货异常
                if (MyConfig.ApplicationConfig.SendDeliverPlanFinishStatisticErrorOnce)
                {
                    DeliverPlanFinishStatisticErrorSender();
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["SendDeliverPlanFinishStatisticErrorOnce"].BoolValue = false;
                    cfg.SaveToFile(fileName);
                }
                //发ECR
                if (MyConfig.ApplicationConfig.SendECRECNOnce)
                {
                    ECRECNProgressSender();
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["SendECRECNOnce"].BoolValue = false;
                    cfg.SaveToFile(fileName);
                }
                //发入库跟催表
                if (MyConfig.ApplicationConfig.SendProduceUnFinishInStockOnce)
                {
                    SendProduceUnFinishInStockEmailLoder();
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["SendProduceUnFinishInStockOnce"].BoolValue = false;
                    cfg.SaveToFile(fileName);
                }
                // 重算工单成本
                if (MyConfig.ApplicationConfig.ProduceEstimateOnceTimes > 0)
                {
                    if (CompanyType == "MY")
                    {
                        int xDays = -MyConfig.ApplicationConfig.ProduceEstimateOnceTimes;
                        ProduceEstimateLastRunTime = ProduceEstimateLastRunTime.AddDays(xDays);
                        ProduceEstimateLoader();
                    }
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["ProduceEstimateOnceTimes"].IntValue = 0;
                    cfg.SaveToFile(fileName);
                }
                // 临时当班发达成率
                if (MyConfig.ApplicationConfig.SendProdPlanEmailOnce)
                {
                    SendProdPlanEmailLoader(NowTime);  //定時發送排程和達成率。
                    MyRecord.Say("发送排程完成。");
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["SendProdPlanEmailOnce"].BoolValue = false;
                    cfg.SaveToFile(fileName);
                }

                // 临时发不合格判定统计表
                if (MyConfig.ApplicationConfig.SendIPQCListStaticOnce)
                {
                    SendIPQCListStaticLoader();// 发送不合格判定统计表
                    MyRecord.Say("临时发送不合格判定统计表。");
                    string fileName = string.Format("{0}\\Config.ini", System.Windows.Forms.Application.StartupPath);
                    SharpConfig.Configuration cfg = SharpConfig.Configuration.LoadFromFile(fileName);
                    cfg["Application"]["SendIPQCListStaticOnce"].BoolValue = false;
                    cfg.SaveToFile(fileName);
                }

                //计时器归零
                if (h == 0 && m == 0 & s == 0)
                {   ///每天0点对表，计时器归零。
                    ProduceFeedBackLastRunTime = CheckStockStartTime = CalculatePruchaseStartTime = SendIPQCListStaticStartTime = KickOutSPUserStartTime = CheckPlanConfirmStartTime = CalculateOrderStartTime = DateTime.Now;
                    MyRecord.Say("重置计数器");
                }
                ///每小时候审核一次仓库单据
                if ((NowTime - CheckStockStartTime).TotalHours > 1.2)
                {
                    MyRecord.Say("每小时审核一次仓库单据，准备开启线程。");
                    if (!CheckStockRecordRunning)
                    {
                        CheckStockStartTime = NowTime;
                        CheckStockRecordRunning = true;
                        CheckStockRecordLoder();   //定時審核出入庫單。
                        Thread.Sleep(200);
                    }
                    else
                    {
                        MyRecord.Say("上一个正在运行中。");
                    }
                    MyRecord.Say("开启完毕");
                }
                //每隔1.4个小时计算一次采购数量
                if ((NowTime - CalculatePruchaseStartTime).TotalHours > 1.4)
                {
                    CalculatePruchaseStartTime = NowTime;
                    PurchaseCalculateLoader();
                    Thread.Sleep(200);
                }
                //每隔42分钟计算一次送货数到订单
                if ((NowTime - CalculateOrderStartTime).TotalHours > 0.7)
                {
                    CalculateOrderStartTime = NowTime;
                    OrderCalculateLoader();
                    Thread.Sleep(200);
                }
                //每隔9分钟跑一次审核30分钟撤销的排程
                if ((NowTime - CheckPlanConfirmStartTime).TotalMinutes > 13)
                {
                    CheckPlanConfirmStartTime = NowTime;
                    ConfirmPlanHalfHour();
                    Thread.Sleep(200);
                }
                //每隔11分钟跑一次踢出特殊权限。
                if ((NowTime - KickOutSPUserStartTime).TotalMinutes > 11)
                {
                    KickOutSPUserStartTime = NowTime;
                    KickOutSpecialHalfHour();
                    Thread.Sleep(200);
                }
                //每隔4H一次发不良判定统计表
                if ((NowTime - SendIPQCListStaticStartTime).TotalHours > 4)
                {
                    if ((NowTime.DayOfWeek != DayOfWeek.Sunday && NowTime.DayOfWeek != DayOfWeek.Monday) || 
                        (NowTime.DayOfWeek == DayOfWeek.Sunday && NowTime.Hour < 9) ||
                        (NowTime.DayOfWeek == DayOfWeek.Monday && NowTime.Hour > 8))
                    {
                        SendIPQCListStaticStartTime = NowTime;
                        SendIPQCListStaticLoader(); //发送不合格判定统计表
                        Thread.Sleep(200);
                    }
                }

                if (h == 5 || h == 8 || h == 10 || h == 11 || h == 14 || h == 17 || h == 19 || h == 22)  //保存达成率和计算看板
                {
                    if (m == 9 && s == 14)
                    {
                        MyRecord.Say("开启保存达成率线程");
                        ProdPlanForSaveLoader(NowTime);//保存达成率到月报表。
                    }
                    else if (m == 33 && s == 17)
                    {
                        MyRecord.Say("每天定時審核纪律单");
                        WorkspaceInspectCheckLoader(); //审核纪律单。
                        Thread.Sleep(5000);   //等五秒。
                        MyRecord.Say("开启计算看板");
                        KanbanRecorderLoader();
                    }
                }
                if ((h == 6 || h == 9 || h == 13 || h == 15 || h == 19 || h == 22)) //计算OEE
                {
                    if (m == 15 && s == 15)
                    {
                        MyRecord.Say("开启计算OEE线程");
                        OEE_ForSaveLoader(NowTime);
                    }
                }
                if ((h == 8 && m == 45 && s == 15) || (h == 21 && m == 05 && s == 15)) //审核排程
                {
                    MyRecord.Say("每天定时審核排程");
                    ConfirmProcessPlanLoader(); //每天定时审核单据，先审核单据。
                }
                if ((h == 9 || h == 22) && m == 5 && s == 15) //发达成率
                {
                    DateTime xDate;
                    if (h == 9)
                        xDate = NowTime.AddDays(-1).Date.AddHours(22);
                    else
                        xDate = NowTime.AddDays(-1).Date.AddHours(10);
                    SendProdPlanEmailLoader(xDate);  //定時發送排程和達成率。
                    MyRecord.Say("发送排程完成。");
                }
                else if ((h == 11 || h == 23) && m == 09 && s == 25) //自動計算完工數
                {
                    ProduceFeedBackLoder();
                    Thread.Sleep(500);
                }
                else if (h == 2 && m == 3 && s == 30)   //计算工单成本
                {
                    if (CompanyType == "MY")
                    {
                        ProduceEstimateLoader();
                        Thread.Sleep(1000);
                    }
                }
                else if (h == 4 && m == 2 && s == 1) //自动计算库存的最后出库日期，平均周转天数，反馈入库时间到出库表
                {
                    StockCalculateLoader();
                }
                else if (((h == 5 || h == 17) && m == 35 && s == 10)) //清理排程
                {
                    if (!PlanRecordCleanRuning) PlanRecordCleanLoder(); ///每天清理排程
                }
                else if (h == 20 && m == 15 && s == 8) //清理LOG
                {
                    //if (CompanyType == "MY") ImportEmployeeLoader();
                    LogingFileCleanLoader();
                }

                if (d == 2 && h == 8 && m == 3 && s == 17)  //每月1日8点发送上个月的达成率和纪律
                {
                    SendKanbanEmail_Performance(NowTime);
                    SendKanbanEmail_Inspect(NowTime);
                }

                if (w != DayOfWeek.Sunday)  //工作日发
                {
                    if (h == 7)
                    {
                        if (m == 1 && s == 15)       //发送每日出货计划量
                        {
                            DeliverPlanNumbSumSender();
                        }
                        else if (m == 5 && s == 15)  //发送不良超100%
                        {
                            RejectSendMailLoader();  ///发送不良率超过100%的列表，每天早6点发送。
                        }
                        else if (m == 10 && s == 15) //发送未审核工单
                        {
                            SendProduceDayReportLoder(); //每天发送未审核工单，每天早7点发送。
                        }
                        else if (m == 12 && s == 15) //未结单工作日发送。
                        {
                            SendProduceUnFinishEmailLoder();
                        }
                        else if (m == 15 && s == 15) //每天发送8D报告跟踪表
                        {
                            ProblemSolving8DReportLoder();
                        }
                        else if (m == 20 && s == 30) //发送未判定和未退料
                        {
                            SendIPQCAndScrapBackEmailLoder();
                        }
                        else if (m == 35 && s == 7)  //发送入库跟催表。
                        {
                            MyRecord.Say("发送入库跟催表，时间循环启动。");
                            SendProduceUnFinishInStockEmailLoder();
                        }
                        else if (m == 55 && s == 10) //发送昨天工单差异数
                        {
                            SendProduceDiffNumbEmailLoder2();
                        }
                    }
                    else if (h == 8)  //发送品保纪律稽核
                    {
                        if (m == 15 && s == 15)
                        {
                            if (CompanyType == "MY")
                            {
                                MyRecord.Say("发送OQC");
                                SendOQCEmailLoder();
                            }
                        }

                        if (m == 30 && s == 15)
                        {
                            if (CompanyType == "MY")
                            {
                                MyRecord.Say("发送ECR/ECN");
                                ECRECNProgressSender();
                            }
                        }

                        if (m == 50 && s == 10)
                        {
                            if (CompanyType == "MY")
                            {
                                MyRecord.Say("发送品保稽核");
                                SendWorkspaceInspectionLoder();
                            }
                        }

                    }
                    else if (h == 11 && m == 55 && s == 51) //自动计算“三日出货计划异常”异常项目发送邮件。
                    {
                        DeliverPlanFinishStatisticErrorSender();
                    }
                    else if (h == 12 && m == 35 && s == 55)  //自动发送未结束维修申请//2周未审核的作废
                    {
                        MachineRepairReportLoder();
                    }
                    else if (h == 13 && m == 45 && s == 9) //发送昨日结单的成本
                    {
                        if (CompanyType == "MY")
                        {
                            SendProduceEstimateLoder();
                        }
                    }
                    //else if (h == 22 && m == 55 && s == 52) //发送工单差异数
                    //{
                    //    SendProduceDiffNumbEmailLoder();   ///每天零点发送昨天工单差异数
                    //}
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        #region 記錄行爲
        protected override void OnStop()
        {
            MyRecord.Say("服务被停止。");
            MainTimer.Stop();
            StopAll();
        }

        protected override void OnContinue()
        {
            MyRecord.Say("服务被重新启动。");
            MainTimer.Start();
        }

        protected override void OnPause()
        {
            MyRecord.Say("服务被暂停。");
            MainTimer.Stop();
            StopAll();
        }

        protected override void OnShutdown()
        {
            MyRecord.Say("服务被关闭");
            MainTimer.Stop();
            StopAll();
        }
        #endregion

        #region 结束所有

        bool _StopAll = false;

        void StopAll()
        {
            MyRecord.Say("关闭线程。");
            _StopAll = true;
            Thread.Sleep(2000);
        }

        #endregion

        #region 审核单据

        protected internal bool RecordCheck(string mainTableName, string statusTableName, int CurrentID, string CurrentRdsNO)
        {
            try
            {
                string SQL = @"SET NOCOUNT ON
						   if Exists(select * from syscolumns where LOWER(name)='checker' and id=object_id('" + mainTableName + @"'))
							Begin
								Update [" + mainTableName + @"] Set Checker='自動審核',CheckDate=GetDate(),Status=1 Where _ID=@zbid
							End
						   Insert Into [" + statusTableName + @"]
								 (zbid,Author,[state],memo,rdsno,type,typeid,CheckIn,CheckOut)
						   Values(@zbid,'自動審核',1,@memo,@rdsno,'審核',2,1,1)
						   SET NOCOUNT OFF ";
                MyData.MyDataParameter[] mp = new MyData.MyDataParameter[3]
				{
					new MyData.MyDataParameter("@zbid", CurrentID, MyData.MyDataParameter.MyDataType.Int),
					new MyData.MyDataParameter("@memo", null),
					new MyData.MyDataParameter("@rdsno", CurrentRdsNO)
				};
                MyData.MyCommand mc = new MyData.MyCommand();
                return mc.Execute(SQL, mp);
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
                return false;
            }
        }

        #endregion

        #region 审核出入库单

        public bool CheckStockRecordRunning = false;

        void CheckStockRecordLoder()
        {
            MyRecord.Say("开启定时审核出入库线程..........");
            Thread t = new Thread(CheckStockRecord);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("定时审核出入库线程已经开启。");
        }

        void CheckStockRecord()
        {
            DateTime rxtime = DateTime.Now;
            while (ProduceFeedBackRuning)
            {
                try
                {
                    if ((DateTime.Now - rxtime).TotalMinutes > 20) return;
                    Thread.Sleep(5000);
                }
                finally
                {
                    CheckStockRecordRunning = false;
                }
            };
            string SQL = "";
            MyRecord.Say("---------------------启动定时审核出入库单据。------------------------------");
            DateTime NowTime = DateTime.Now;
            DateTime StopTime = NowTime.AddHours(0 - CheckStockTimers);
            MyRecord.Say(string.Format("审核截止时间：{0:yy/MM/dd HH:mm}", StopTime));
            CheckStockRecordRunning = true;
            try
            {
                #region 審核產成品入庫----新版
                try
                {
                    MyRecord.Say("审核产成品入库单——新ERP系统审核");
                    SQL = @"Select * from stProduceStock Where InputDate < @InputEnd And isNull(Status,0)=0 And isono='NEWERP'";
                    MyData.MyDataTable mTableProduceInStock = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableProduceInStock != null && mTableProduceInStock.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableProduceInStock.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableProduceInStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableProduceInStock.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["_ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableProduceInStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            if (!RecordCheck("stProduceStock", "_WH_ProduceEntryWarehouse_StatusRecorder", CurID, RdsNo))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableProduceInStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核产成品入库单——新ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核其他入库----新版
                try
                {
                    MyRecord.Say("审核其他入库单——新ERP系统审核");
                    SQL = @"Select * from stOtherStock Where InputDate < @InputEnd And isNull(Status,0)=0 And otherno='NEWERP'";
                    MyData.MyDataTable mTableOtherInStock = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableOtherInStock != null && mTableOtherInStock.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableOtherInStock.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableOtherInStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableOtherInStock.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["_ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableOtherInStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            if (!RecordCheck("stOtherStock", "_WH_OtherStockIn_StatusRecorder", CurID, RdsNo))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableOtherInStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核其他入库单——新ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核生产领料----新版
                try
                {
                    MyRecord.Say("审核生产领料单——新ERP系统审核");
                    SQL = @"Select * from stProduceOut Where InputDate < @InputEnd And isNull(Status,0)=0 And ISONO='NEWERP'";
                    MyData.MyDataTable mTableOtherInStock = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableOtherInStock != null && mTableOtherInStock.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableOtherInStock.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableOtherInStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableOtherInStock.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["_ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableOtherInStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            if (!RecordCheck("stProduceOut", "_WH_ProducePicking_StatusRecorder", CurID, RdsNo))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableOtherInStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核其他入库单——新ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核產成品入庫---旧版
                try
                {
                    MyRecord.Say("审核产成品入库单——旧ERP系统审核");
                    SQL = @"Select * from stProduceStock Where InputDate < @InputEnd And isNull(Status,0)=0 And isNull(isono,'')=''";
                    MyData.MyDataTable mTableProduceInStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableProduceInStock_OLD != null && mTableProduceInStock_OLD.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableProduceInStock_OLD.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableProduceInStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableProduceInStock_OLD.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["_ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableProduceInStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            string UpdateOLD_SQL = @"Update stProduceStock Set CheckDate=GetDate(),Checker='自動審核',Status=1 Where RdsNo=@RdsNo";
                            MyData.MyCommand mc = new MyData.MyCommand();
                            if (!mc.Execute(UpdateOLD_SQL, new MyData.MyDataParameter("@rdsno", RdsNo)))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableProduceInStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核产成品入库单——旧ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核其他入庫---旧版
                try
                {
                    MyRecord.Say("审核其他入库单——旧ERP系统审核");
                    SQL = @"Select * from stOtherStock Where InputDate < @InputEnd And isNull(Status,0)=0 And isNull(otherno,'')<>'NEWERP'";
                    MyData.MyDataTable mTableOtherInStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableOtherInStock_OLD != null && mTableOtherInStock_OLD.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableOtherInStock_OLD.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableOtherInStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableOtherInStock_OLD.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["_ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableOtherInStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            string UpdateOLD_SQL = @"Update stOtherStock Set CheckDate=GetDate(),Checker2='自動審核',Status=1 Where RdsNo=@RdsNo";
                            MyData.MyCommand mc = new MyData.MyCommand();
                            if (!mc.Execute(UpdateOLD_SQL, new MyData.MyDataParameter("@rdsno", RdsNo)))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableOtherInStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核其他入库单——旧ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核采购入庫---旧版
                try
                {
                    MyRecord.Say("审核产成品入库单——旧ERP系统审核");
                    SQL = @"Select * from coPurchStock Where InputDate < @InputEnd And isNull(Status,0)=0 And ISONO <> 'NEWERP'";
                    MyData.MyDataTable mTablePurchInStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTablePurchInStock_OLD != null && mTablePurchInStock_OLD.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTablePurchInStock_OLD.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableProduceInStockCount = 1;
                        foreach (MyData.MyDataRow r in mTablePurchInStock_OLD.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableProduceInStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            string UpdateOLD_SQL = @"Update coPurchStock Set CheckDate=GetDate(),Checker='自動審核',Status=1 Where RdsNo=@RdsNo";
                            MyData.MyCommand mc = new MyData.MyCommand();
                            if (!mc.Execute(UpdateOLD_SQL, new MyData.MyDataParameter("@rdsno", RdsNo)))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableProduceInStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核产成品入库单——旧ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核采购入库----新版
                try
                {
                    MyRecord.Say("审核采购入库——新ERP系统审核");
                    SQL = @"Select * from coPurchStock Where InputDate < @InputEnd And isNull(Status,0)=0 And ISONO='NEWERP'";
                    MyData.MyDataTable mTableOtherInStock = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableOtherInStock != null && mTableOtherInStock.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableOtherInStock.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableOtherInStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableOtherInStock.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["_ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableOtherInStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            if (!RecordCheck("coPurchStock", "_WH_PurchaseEntryWarehouse_StatusRecorder", CurID, RdsNo))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableOtherInStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核其他入库单——新ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核生产领料---旧版
                try
                {
                    MyRecord.Say("审核生产领料——旧ERP系统审核");
                    SQL = @"Select * from stProduceOut Where InputDate < @InputEnd And isNull(Status,0)=0 And ISONO is Null ";
                    MyData.MyDataTable mTableProduceOutStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableProduceOutStock_OLD != null && mTableProduceOutStock_OLD.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableProduceOutStock_OLD.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableProduceOutStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableProduceOutStock_OLD.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableProduceOutStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            string UpdateOLD_SQL = @"Update stProduceOut Set CheckDate=GetDate(),Checker2='自動審核',Status=1 Where RdsNo=@RdsNo";
                            MyData.MyCommand mc = new MyData.MyCommand();
                            if (!mc.Execute(UpdateOLD_SQL, new MyData.MyDataParameter("@rdsno", RdsNo)))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableProduceOutStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核生产领料——旧ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核其他出库---旧版
                try
                {
                    MyRecord.Say("审核其他领料——旧ERP系统审核");
                    SQL = @"Select * from stOtherOut Where InputDate < @InputEnd And isNull(Status,0)=0 And handno is Null ";
                    MyData.MyDataTable mTableOthereOutStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableOthereOutStock_OLD != null && mTableOthereOutStock_OLD.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableOthereOutStock_OLD.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableOtherOutStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableOthereOutStock_OLD.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableOtherOutStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            string UpdateOLD_SQL = @"Update stOtherOut Set CheckDate=GetDate(),Checker2='自動審核',Status=1 Where RdsNo=@RdsNo";
                            MyData.MyCommand mc = new MyData.MyCommand();
                            if (!mc.Execute(UpdateOLD_SQL, new MyData.MyDataParameter("@rdsno", RdsNo)))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableOtherOutStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核生产领料——旧ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核其他领料----新版
                try
                {
                    MyRecord.Say("审核其他入库单——新ERP系统审核");
                    SQL = @"Select * from stOtherOut Where InputDate < @InputEnd And isNull(Status,0)=0 And handno='NEWERP'";
                    MyData.MyDataTable mTableOtherInStock = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableOtherInStock != null && mTableOtherInStock.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableOtherInStock.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableOtherInStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableOtherInStock.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["_ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableOtherInStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            if (!RecordCheck("stOtherOut", "_WH_OtherStockOut_StatusRecorder", CurID, RdsNo))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableOtherInStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核其他入库单——新ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                #region 審核送货单----新版
                if (CompanyType != "MT" || NowTime.Date == NowTime.Date.AddMonths(1).AddDays(-NowTime.Day).Date)
                {
                    try
                    {
                        MyRecord.Say("审核送货单——新ERP系统审核");
                        SQL = @"Select * from coShip Where InputDate < @InputEnd And isNull(Status,0)=0 ";
                        MyData.MyDataTable mTableOtherInStock = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                        if (_StopAll) return;
                        if (mTableOtherInStock != null && mTableOtherInStock.MyRows.Count > 0)
                        {
                            string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableOtherInStock.MyRows.Count);
                            MyRecord.Say(memo);
                            int mTableOtherInStockCount = 1;
                            foreach (MyData.MyDataRow r in mTableOtherInStock.MyRows)
                            {
                                if (_StopAll) return;
                                int CurID = Convert.ToInt32(r["_ID"]);
                                string RdsNo = Convert.ToString(r["RdsNo"]);
                                memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableOtherInStockCount, RdsNo, CurID, r["InputDate"]);
                                MyRecord.Say(memo);
                                if (!RecordCheck("coShip", "_WH_DeiliverNote_StatusRecorder", CurID, RdsNo))
                                {
                                    MyRecord.Say("审核错误！");
                                }
                                mTableOtherInStockCount++;
                            }
                        }
                        else
                        {
                            MyRecord.Say("没有获取到任何内容");
                        }
                        MyRecord.Say("审核送货单——新ERP系统审核，完成");
                    }
                    catch (Exception ex)
                    {
                        MyRecord.Say(ex);
                    }
                }
                else
                {
                    MyRecord.Say("送货单——不审核。");
                }
                #endregion

                #region 審核库存调整单----新版&旧版
                try
                {
                    MyRecord.Say("审核库存调整单——新ERP系统审核");
                    SQL = @"Select * from stAdjustStock Where InputDate < @InputEnd And isNull(Status,0)=0";
                    MyData.MyDataTable mTableAdjustStock = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                    if (_StopAll) return;
                    if (mTableAdjustStock != null && mTableAdjustStock.MyRows.Count > 0)
                    {
                        string memo = string.Format("读取了{0}条记录，下面开始审核....", mTableAdjustStock.MyRows.Count);
                        MyRecord.Say(memo);
                        int mTableAdjustStockCount = 1;
                        foreach (MyData.MyDataRow r in mTableAdjustStock.MyRows)
                        {
                            if (_StopAll) return;
                            int CurID = Convert.ToInt32(r["_ID"]);
                            string RdsNo = Convert.ToString(r["RdsNo"]);
                            memo = string.Format("审核第{0}条，入库单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableAdjustStockCount, RdsNo, CurID, r["InputDate"]);
                            MyRecord.Say(memo);
                            if (!RecordCheck("stAdjustStock", "_WH_AdjustInventory_StatusRecorder", CurID, RdsNo))
                            {
                                MyRecord.Say("审核错误！");
                            }
                            mTableAdjustStockCount++;
                        }
                    }
                    else
                    {
                        MyRecord.Say("没有获取到任何内容");
                    }
                    MyRecord.Say("审核库存调整单——新旧ERP系统审核，完成");
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                #endregion

                Thread.Sleep(1000);
            }
            finally
            {
                CheckStockRecordRunning = false;
                MyRecord.Say("-----------------------审核完成。-------------------------------");
            }
        }

        #endregion

        #region 不良率100发邮件
        void RejectSendMailLoader()
        {
            MyRecord.Say("开启定时发送不良率100%线程..........");
            Thread t = new Thread(RejectSendMail);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("定时发送不良率100%线程应启动。");
        }

        void RejectSendMail()
        {
            try
            {
                MyRecord.Say("------------------开始定时发送不良率100%报表----------------------------");
                MyRecord.Say("从数据库搜寻内容");
                string SQL = @"
select
a.ProcessID as ProcessCode,a.ProductCode,a.Numb2,a.Operator,a.PartID,a.StartTime,a.EndTime,a.Remark2,a.AccNoteRdsNo,a.RptDate,a.ProduceNo,Remark,
ProcessName=(Select Name from moProcedure bp Where bp.Code=a.ProcessID),
MachineName=(Select Name from moMachine bm Where bm.Code=a.MachinID),
DepartmentName=(Select pp.Name from pbDept pp,moMachine pm Where pp.[_id]=pm.DepartmentID And pm.Code=a.MachinID),
ProductName = (Select Name from pbProduct Where Code=a.ProductCode),
ProduceNumb = (Select PNumb From moProduce Where RdsNo=a.ProduceNo),
PNumb=(Select mb.ReqNumb from moProduce ma,moProdProcedure mb Where ma.id=mb.zbid And ma.RdsNo=a.ProduceNo And mb.id=a.PrdID),
a.NAColNumb,
Inputer
from prodDailyReport a Where a.RptDate Between @DateBegin And @DateEnd
								And isNull(a.Numb1,0) = 0 AND isNull(a.Numb2,0) > 0
								And a.ProcessID<>'2000'
								And isNull(a.Reject,0) = 0
Order by a.ProcessID,a.MachinID
";
                DateTime NowTime = DateTime.Now;
                MyData.MyDataParameter[] myps = new MyData.MyDataParameter[]
			{
				new MyData.MyDataParameter("@DateBegin",NowTime.AddDays(-1).Date, MyData.MyDataParameter.MyDataType.DateTime),
				new MyData.MyDataParameter("@DateEnd",NowTime.Date , MyData.MyDataParameter.MyDataType.DateTime)
			};
                MyRecord.Say(string.Format("@DateBegin = {0}，@DateBegin = {1}", NowTime.AddDays(-1).Date, NowTime.Date));
                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#e6f3af #ffffff>
<DIV><FONT size=3 face=PMingLiU>{3}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {0:yy/MM/dd}完工单中有如下内容不良率达到100%（完工单上只有不良品）。</FONT></DIV>
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
	机台
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	料号(部件)
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单数(PCS)
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工序应产数
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	良品数
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	不良數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	完工單/验收入库单
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	输入人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	机长
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	备注
	</TD>
	</TR>
	{1}
</TBODY></TABLE></FONT>
</DIV>
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，切勿在此郵件上直接回復。</STRONG></FONT></DIV>
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
	</TR>
";
                string brs = "";
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, myps))
                {
                    if (md != null && md.MyRows.Count > 0)
                    {
                        MyRecord.Say("生成表格");
                        foreach (var ri in md.MyRows)
                        {
                            brs += string.Format(br, Convert.ToString(ri["DepartmentName"]),
                                                    Convert.ToString(ri["ProcessName"]),
                                                    Convert.ToString(ri["MachineName"]),
                                                    Convert.ToString(ri["ProduceNo"]),
                                                    string.Format("{0} {1}", Convert.ToString(ri["ProductName"]), Convert.ToString(ri["PartID"])),
                                                    Convert.ToInt32(ri["ProduceNumb"]),
                                                    Convert.ToInt32(ri["PNumb"]),
                                                    0,
                                                    string.Format("{0}x{1}", Convert.ToInt32(ri["Numb2"]), Convert.ToInt32(ri["NAColNumb"])),
                                                    Convert.ToString(ri["Remark2"]).Length <= 0 ? Convert.ToString(ri["AccNoteRdsNo"]) : Convert.ToString(ri["Remark2"]),
                                                    Convert.ToString(ri["Inputer"]),
                                                    Convert.ToString(ri["Operator"]),
                                                    Convert.ToString(ri["Remark"])
                                                    );
                        }
                        MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", md.Rows.Count));
                        MyRecord.Say("创建SendMail。");
                        MyBase.SendMail sm = new MyBase.SendMail();
                        MyRecord.Say("加载邮件内容。");
                        sm.MailBodyText = string.Format(body, NowTime.AddDays(-1).Date, brs, DateTime.Now, MyBase.CompanyTitle);
                        sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}完工单不良率100%统计表", NowTime.AddDays(-1).Date, MyBase.CompanyTitle));
                        //string mailto = ConfigurationManager.AppSettings["RejectMailTo"], mailcc = ConfigurationManager.AppSettings["RejectMailCC"];
                        //MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        //sm.MailTo = mailto;
                        //sm.MailCC = mailcc;
                        ////sm.MailTo = "my80@my.imedia.com.tw";

                        MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("Reject");
                        MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                        sm.MailTo = mAddress.MailTo;
                        sm.MailCC = mAddress.MailCC;
                        //sm.MailTo = "my80@my.imedia.com.tw";

                        MyRecord.Say("发送邮件。");
                        sm.SendOut();
                        MyRecord.Say("已经发送。");
                    }
                    else
                    {
                        MyRecord.Say("没有找到资料。");
                    }
                }
                MyRecord.Say("------------------发送完成----------------------------");
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }
        #endregion

        #region 发送3天前没有结案的工单

        void SendProduceUnFinishEmailLoder()
        {
            MyRecord.Say("开启定时发送10天未结单进程..........");
            Thread t = new Thread(SendProduceUnFinishEmail);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启定时发送10天未结单进程完成。");
        }

        void SendProduceUnFinishEmail()
        {
            try
            {
                MyRecord.Say("-----------------开启定时发送未结工单-------------------------");
                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{5}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意以下未结单，详情请查看附档。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; （说明：工单未结是指—工单需要入库且入库完成没有打勾的工单。）</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	项目
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单数目
	</TD>
	</TR>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	三天前（{0:yy/MM/dd}前含当日）审核，交期小于今日的未结工单
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{1}
	</TD>
	</TR>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	七天前（{2:yy/MM/dd}前含当日）审核的未结工单
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{3}
	</TD>
	</TR>
</TBODY></TABLE></FONT>
</DIV>
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{4:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");

                string SQL1 = @"
SELECT [a].[rdsno],[a].[orderno],[a].[custid],[a].[code],[a].[pdeliver],[a].[inputdate],[a].[inputer],[a].[checkdate],[a].[checker],[a].[remark],[a].[finishdate],
	   [a].[property],[a].[FinishRemark],[a].[FinishMan],[a].[SRemark],[a].[RMan],[a].[RTime],[a].[InStockFinishNumb],[c].[name],[c].[Size],[c].[mTypeName],
	   [mpp].[Name] AS [PropertyName],[a].[ProdFinishNumb],[opp].[Name] AS [TypeName],[b1].[Type],[a].[pnumb],[a].[PickedNumb],[a].[PickedNumb2],
	   DATEDIFF(DAY,[a].[checkdate],GETDATE()) AS CheckDateDiff,pnn.[N] AS ProduceTimes,[pnn].[OnProduceNumb] AS ProducingNumb
FROM [moProduce] [a] INNER JOIN [coOrder] [b1] ON[a].[orderno] = [b1].[rdsno]
				   INNER JOIN [coOrderProd] [b2] ON [b1].[id]=[b2].[zbid] AND [a].[pid]=[b2].[id]
				   INNER JOIN [AllMaterialView] [c] ON [a].[code]=[c].[code]
				   INNER JOIN [moProdProperty] [mpp] ON [a].[property] = [mpp].[Code]
				   INNER JOIN [_SD_OrderNote_Type] [opp] ON [b1].[Type]= [opp].[Code]
				   LEFT OUTER JOIN 
				   (SELECT Count(*) as N,[ax].[code],SUM([ax].[pnumb]-[ax].[InStockFinishNumb]) AS OnProduceNumb
					  FROM [moProduce] [ax]
									Where ISNULL([ax].[status],0) > 0 
									AND ((isNull([ax].[StopStock],0)=0 And [ax].[StockDate] is Null) or (isNull([ax].[StopStock],0)=1 and [ax].[finishdate] is Null))
									GROUP BY [ax].[code]) [pnn] ON [pnn].[code] = [a].[code]
WHERE [a].[StockDate] IS NULL
	   AND YEAR([a].[inputdate]) > 2013
	   AND ISNULL([a].[StopStock], 0) = 0
	   AND DATEDIFF(DAY,[a].[checkdate],GETDATE()) > 3 
	   AND DATEDIFF(DAY,[a].[pdeliver],GETDATE()) > 0
ORDER BY DATEDIFF(DAY,[a].[checkdate],GETDATE()) DESC
",
 SQL2 = @"
SELECT [a].[rdsno],[a].[orderno],[a].[custid],[a].[code],[a].[pdeliver],[a].[inputdate],[a].[inputer],[a].[checkdate],[a].[checker],[a].[remark],[a].[finishdate],
	   [a].[property],[a].[FinishRemark],[a].[FinishMan],[a].[SRemark],[a].[RMan],[a].[RTime],[a].[InStockFinishNumb],[c].[name],[c].[Size],[c].[mTypeName],
	   [mpp].[Name] AS [PropertyName],[a].[ProdFinishNumb],[opp].[Name] AS [TypeName],[b1].[Type],[a].[pnumb],[a].[PickedNumb],[a].[PickedNumb2],
	   DATEDIFF(DAY,[a].[checkdate],GETDATE()) AS CheckDateDiff,pnn.[N] AS ProduceTimes,[pnn].[OnProduceNumb] AS ProducingNumb
FROM [moProduce] [a] INNER JOIN [coOrder] [b1] ON[a].[orderno] = [b1].[rdsno]
				   INNER JOIN [coOrderProd] [b2] ON [b1].[id]=[b2].[zbid] AND [a].[pid]=[b2].[id]
				   INNER JOIN [AllMaterialView] [c] ON [a].[code]=[c].[code]
				   INNER JOIN [moProdProperty] [mpp] ON [a].[property] = [mpp].[Code]
				   INNER JOIN [_SD_OrderNote_Type] [opp] ON [b1].[Type]= [opp].[Code]
				   LEFT OUTER JOIN 
				   (SELECT Count(*) as N,[ax].[code],SUM([ax].[pnumb]-[ax].[InStockFinishNumb]) AS OnProduceNumb
					  FROM [moProduce] [ax]
									Where ISNULL([ax].[status],0) > 0 
									AND ((isNull([ax].[StopStock],0)=0 And [ax].[StockDate] is Null) or (isNull([ax].[StopStock],0)=1 and [ax].[finishdate] is Null))
									GROUP BY [ax].[code]) [pnn] ON [pnn].[code] = [a].[code]
WHERE [a].[StockDate] IS NULL
	   AND YEAR([a].[inputdate]) > 2013
	   AND ISNULL([a].[StopStock], 0) = 0
	   AND DATEDIFF(DAY,[a].[checkdate],GETDATE()) > 7
ORDER BY DATEDIFF(DAY,[a].[checkdate],GETDATE()) DESC
";

                DateTime NowTime = DateTime.Now;
                MyRecord.Say("后台计算");
                MyData.MyDataTable md1 = new MyData.MyDataTable(SQL1);
                MyData.MyDataTable md2 = new MyData.MyDataTable(SQL2);

                MyRecord.Say("工单已经获取");

                int ThreedayProduceCount = md1.MyRows.Count, SevendayProduceCount = md2.MyRows.Count;
                MyRecord.Say(string.Format("1表格一共：{0}行，内容已经生成。", md1.Rows.Count));
                MyRecord.Say(string.Format("2表格一共：{0}行，内容已经生成。", md2.Rows.Count));
                MyRecord.Say("创建SendMail。");
                MyBase.SendMail sm = new MyBase.SendMail();
                MyRecord.Say("加载邮件内容。");
                DateTime ThreeDay = NowTime.AddDays(-2).Date, SevenDay = NowTime.AddDays(-6).Date;
                sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, ThreeDay, ThreedayProduceCount, SevenDay, SevendayProduceCount, NowTime, MyBase.CompanyTitle));
                sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_未结工单提醒。", NowTime, MyBase.CompanyTitle));


                HSSFWorkbook wb = new HSSFWorkbook();
                string[] fNames = new string[] { "rdsno", "orderno", "custid", "TypeName", "PropertyName", "code", "NAME", "mTypeName", "pnumb", "PickedNumb2", "ProdFinishNumb", "InStockFinishNumb", "checkdate", "checker", "CheckDateDiff", "ProduceTimes", "ProducingNumb", "pdeliver", "SRemark", "RMan", "RTime", "inputdate", "inputer" };
                string[] fTitles = new string[] { "工单号", "订单号", "客户", "订单类型", "生产性质", "编号", "料号", "类型", "应产数", "领料数", "完工数", "入库数", "审核日期", "审核人", "审核日期至今天数", "此料未结单数目", "此料工单未结数量", "工单交期", "交期备注", "刻交期人", "刻交期时间", "输入时间", "输入人" };


                ExportToExcel(wb, md1, fNames, fTitles, "三天前审核交期过期的未结工单");
                ExportToExcel(wb, md2, fNames, fTitles, "七天前审核的未结工单");

                string tmpFileName = Path.GetTempFileName();
                MyRecord.Say(string.Format("tmpFileName = {0}", tmpFileName));
                Regex rgx = new System.Text.RegularExpressions.Regex(@"(?<=tmp)(.+)(?=\.tmp)");
                string tmpFileNameLast = rgx.Match(tmpFileName).Value;
                MyRecord.Say(string.Format("tmpFileNameLast = {0}", tmpFileNameLast));
                string dName = string.Format("{0}\\TMPExcel", Application.StartupPath);
                if (!Directory.Exists(dName)) Directory.CreateDirectory(dName);
                string fname = string.Format("{0}\\{1}", dName, string.Format("NOTICE_{0:yyyyMMdd}_tmp{1}.xls", NowTime.Date, tmpFileNameLast));
                MyRecord.Say(string.Format("fname = {0}", fname));
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
                wb.Close();


                //string mailto = ConfigurationManager.AppSettings["TenDayMailTo"], mailcc = ConfigurationManager.AppSettings["TenDayMailCC"];
                //MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                //sm.MailTo = mailto; // "my18@my.imedia.com.tw,xang@my.imedia.com.tw,lghua@my.imedia.com.tw,my64@my.imedia.com.tw";
                //sm.MailCC = mailcc; // "jane123@my.imedia.com.tw,lwy@my.imedia.com.tw,my80@my.imedia.com.tw";
                ////sm.MailTo = "my80@my.imedia.com.tw";
                MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("TenDay");
                MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                sm.MailTo = mAddress.MailTo;
                sm.MailCC = mAddress.MailCC;
                //sm.MailTo = "my80@my.imedia.com.tw";
                MyRecord.Say("发送邮件。");
                sm.SendOut();
                sm.mail.Dispose();
                sm = null;
                MyRecord.Say("已经发送。");
                if (File.Exists(fname)) File.Delete(fname);
                MyRecord.Say("已经发送。");

                MyRecord.Say("------------------发送完成----------------------------");
            }
            catch (Exception e)
            {
                MyRecord.Say(e);
            }
        }

        #endregion

        void ExportToExcel(HSSFWorkbook wb, MyData.MyDataTable mds, string[] GridFileds, string[] GridTitles, string SheetCaption)
        {
            if (mds.MyRows.IsNotEmptySet())
            {
                var vListGridRow = from a in mds.MyRows
                                   select a;
                MyRecord.Say(string.Format("输出表格{0}", SheetCaption));
                HSSFSheet st = (HSSFSheet)wb.CreateSheet(SheetCaption);
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
                xCellCaption1.SetCellValue(LCStr(SheetCaption));
                xCellCaption1.CellStyle = cds3;
                st.AddMergedRegion(new CellRangeAddress(1, 1, 0, GridFileds.Length - 1));

                //MyRecord.Say(string.Format("输出表格{0}，表格已经创建，建立表头。", caption));

                HSSFRow rNoteCaption2 = (HSSFRow)st.CreateRow(2);
                for (int i = 0; i < GridFileds.Length; i++)
                {
                    ICell xCellCaption = rNoteCaption2.CreateCell(i);
                    xCellCaption.SetCellValue(LCStr(GridTitles[i]));
                    xCellCaption.CellStyle = cellNomalStyle;
                    //MyRecord.Say(string.Format("输出表格{0}，标题行，{1}列", caption, LCStr(titles[i])));
                }

                MyRecord.Say(string.Format("输出表格{0}，表格已经创建，建立表头完成。", SheetCaption));

                int uu = 3;

                foreach (var item in vListGridRow)
                {
                    if (item.IsNotNull())
                    {
                        HSSFRow rNoteCaption3 = (HSSFRow)st.CreateRow(uu);
                        for (int i = 0; i < GridFileds.Length; i++)
                        {
                            string fieldname = GridFileds[i];
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
                                MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列，没有这个字段。", SheetCaption, uu, fieldname));
                            }
                        }
                        uu++;
                    }
                }
            }
            MyRecord.Say(string.Format("输出表格{0}完成。", SheetCaption));
        }


        #region 发送应入库没入库的内容

        void SendProduceUnFinishInStockEmailLoder()
        {
            MyRecord.Say("开启定时发送入库跟催表..........");
            Thread t = new Thread(SendProduceUnFinishInStockEmail);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启定时发送入库跟催表进程启动。");
        }

        void SendProduceUnFinishInStockEmail()
        {
            try
            {
                MyRecord.Say("-----------------开启定时发送入库跟催表-------------------------");
                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 12pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{2}ERP系统提示您：</FONT></DIV>
{0}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{1:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");
                string bx = @"
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容为12H（{0:MM/dd HH:mm}）前，已输入验收入库单（完工单）但还没输入对应产成品入库单（仓库）的验收入库单列表。</FONT></DIV>
<DIV><FONT size=5 color=#ff0000 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {1}</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	超48H
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	24~48H
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	12~24H
	</TD>
	</TR>
	{2}
</TBODY></TABLE></FONT>
</DIV>
";
                string bx2 = @"
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 12H（{0:MM/dd HH:mm}）前，验收入库单（完工单）都已经入库完毕。</FONT></DIV>
<DIV><FONT size=5 color=#ff0000 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 很棒，请再接再厉</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</FONT>
</DIV>
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
	</TR>
";

                string SQL = @"
	  Set NoCount On
	  Create Table #Tmp_ProdDailyReport(CustCode nVarChar(40),ProductCode nVarChar(200),ProduceRdsNo nVarChar(40),AccNoteRdsno nVarChar(40),RptDate DateTime,Numb1 Numeric(38,6),
										Operator nVarChar(100),InputDate DateTime,Inputer nVarChar(40),Remark nVarChar(4000),
										MachineCode nVarChar(40),ProcessCode nVarChar(40),PrdID Int,
										MachineName nVarChar(100),DepartmentID int
										)
	  Create Index IX_Tmp_ProdDailyReport_1 On #Tmp_ProdDailyReport (ProduceRdsNo,ProductCode,AccNoteRdsNo)

	  Insert Into #Tmp_ProdDailyReport
	  Select a.CustCode,a.ProductCode,a.ProduceNo,a.AccNoteRdsNo,a.RptDate,a.Numb1,a.Operator,a.InputDate,a.Inputer,a.Remark,a.MachinID,a.ProcessID,a.PrdID,b.Name,b.DepartmentID
	  From ProdDailyReport a Inner Join moMachine b On a.MachinID = b.Code
	  Where a.ProcessID <> '9035' And isNull(a.Numb1,0) > 0 And isNull(a.Reject,0) = 0 And Len(isNull(a.AccNoteRdsNo,''))>0
		And a.RptDate Between @SBDate And @SEDate

	  Create Table #Tmp_Result(CustCode nVarChar(40),ProductCode nVarChar(200),ProduceRdsNo nVarChar(40),AccNoteRdsno nVarChar(40),RptDate DateTime,Numb1 Numeric(38,6),
							   Operator nVarChar(100),InputDate DateTime,Inputer nVarChar(40),Remark nVarChar(4000),MachineCode nVarChar(40),ProcessCode nVarChar(40),PrdID Int,
							   ProcessName nVarChar(100),MachineName nVarChar(100),FullSortID nVarChar(50),DepartmentName nVarChar(80),DepartmentID int,
							   ProductName nVarChar(400),Description nVarChar(400),DeclaretionName nVarChar(400),CustProductCode nVarChar(400),
							   PropertyID Int,StockLocationType nVarChar(20),Status Int,StockStatus Int,
							   InStockNumb Numeric(38,6),StockRdsNo nVarChar(40),StockDate DateTime,StockInputer nVarChar(40),stManger nVarChar(40),IQC nVarChar(40),
							   ProduceStatusName nVarChar(80),ProducepropertyName nVarChar(80),StockLocationTypeName nVarChar(80),ProduceStockStatusName nVarChar(80),
							   FinishMan nVarChar(40),FinishRemark nVarchar(4000),finishdate DateTime,ProduceID int)

	  Insert Into #Tmp_Result(CustCode,ProductCode,ProduceRdsNo,AccNoteRdsno,RptDate,Numb1,Operator,InputDate,Inputer,Remark,MachineCode,ProcessCode,PrdID,DepartmentID,PropertyID,StockLocationType,Status,StockStatus,MachineName,FinishMan,FinishRemark,finishdate,ProduceID)
	  Select a.CustCode,a.ProductCode,ProduceRdsNo,a.AccNoteRdsno,a.RptDate,a.Numb1,a.Operator,a.InputDate,a.Inputer,a.Remark,a.MachineCode,a.ProcessCode,a.PrdID,a.DepartmentID,b.property,b.StockLocationType,b.status,b.StockStatus,a.MachineName,b.FinishMan,b.FinishRemark,b.finishdate,b.id
	  from #Tmp_ProdDailyReport a Inner Join moProduce b On a.ProduceRdsNo = b.RdsNo And a.PrdID = b.LastProcessID
	  Where isNull(b.StopStock,0)=0 And isNull(finalized,0)=0 And b.status > 0

	  Update a Set a.ProduceStatusName = b.Name From [#Tmp_Result] a,[_SY_Status] b Where b.StatusID = isNull(a.Status,0) And b.Type = 'PO'
	  Update a Set a.ProduceStockStatusName = b.Name From [#Tmp_Result] a,[_SY_Status] b Where b.StatusID = isNull(a.StockStatus,0) And b.Type = 'POStatus'
	  Update a Set a.ProducepropertyName = b.Name From [#Tmp_Result] a,[_SY_Status] b Where b.StatusID = isNull(a.PropertyID,0) And b.Type = 'POProperty'
	  Update a Set a.StockLocationTypeName = b.Name From [#Tmp_Result] a,[_BS_Location_Type] b Where a.StockLocationType = b.code
	  Update a Set a.ProcessName = b.Name From [#Tmp_Result] a,moProcedure b Where a.ProcessCode = b.code
	  Update a Set a.DepartmentName = b.name,a.FullSortID = b.FullSortID From [#Tmp_Result] a,pbDept b Where a.DepartmentID = b.[_ID]
	  Update a Set a.StockRdsNo = b.rdsno,a.StockDate = b.StockDate,a.InStockNumb = b.Number,a.StockInputer = b.inputer,a.stManger = b.stManger,a.IQC = b.IQC
			  From [#Tmp_Result] a,stProduceStock b Where a.AccNoteRdsno = b.AccNoteRdsNo
	  Update a Set a.ProductName = b.Name,a.Description = b.Description,a.CustProductCode = b.mCustNO,a.DeclaretionName = b.DeclaretionName
			  From [#Tmp_Result] a,pbProduct b Where a.ProductCode = b.Code

	  Select *,
			InStockTimeSpan = Ceiling(DateDiff(HH,RptDate,isNull(StockDate,GetDate())))
	   from #Tmp_Result
	  Where isNull(InStockNumb,0) <= 0
	  Order by FullSortID,ProcessCode,MachineCode,ProductCode,ProduceRdsNo,AccNoteRdsno
	  Drop Table #Tmp_Result
	  Drop Table #Tmp_ProdDailyReport
	  Set NoCount Off
";
                DateTime NowTime = DateTime.Now,
                         xeTime = NowTime.AddDays(-1).DayOfWeek == DayOfWeek.Sunday ? NowTime.AddDays(-2).Date.AddHours(20) : NowTime.AddDays(-1).Date.AddHours(20),
                         xbTime = CompanyType == "MD" ? xeTime.AddYears(-1) < DateTime.Parse("2017-06-01") ? DateTime.Parse("2017-06-01") : xeTime.AddYears(-1) : xeTime.AddMonths(-10);
                string brs = "";
                MyRecord.Say("后台计算开启。");
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL,
                           new MyData.MyDataParameter("@SEDate", xeTime, MyData.MyDataParameter.MyDataType.DateTime),
                           new MyData.MyDataParameter("@SBDate", xbTime, MyData.MyDataParameter.MyDataType.DateTime)))
                {
                    MyRecord.Say("完工单已经获取");
                    if (md != null && md.MyRows.Count > 0)
                    {
                        var vx = from a in md.MyRows
                                 group a by a.Value("DepartmentName") into g
                                 orderby g.FirstOrDefault().Value("FullSortID")
                                 select new
                                 {
                                     Name = g.Key,
                                     Count1 = g.Where(x => x.IntValue("InStockTimeSpan") > 48).Count(),
                                     Count2 = g.Where(x => x.IntValue("InStockTimeSpan") > 24 && x.IntValue("InStockTimeSpan") < 49).Count(),
                                     Count3 = g.Where(x => x.IntValue("InStockTimeSpan") > 0 && x.IntValue("InStockTimeSpan") < 25).Count()
                                 };
                        foreach (var b in vx)
                        {
                            brs += string.Format(br, b.Name, b.Count1 == 0 ? "" : string.Format("{0}笔", b.Count1),
                                                             b.Count2 == 0 ? "" : string.Format("{0}笔", b.Count2),
                                                             b.Count3 == 0 ? "" : string.Format("{0}笔", b.Count3));
                        }
                        MyRecord.Say("创建SendMail。");
                        MyBase.SendMail sm = new MyBase.SendMail();
                        MyRecord.Say("加载邮件内容。");
                        string stSumMemo = string.Format("合计：{0}笔，详细内容请看附档。", md.MyRows.Count);

                        sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, string.Format(bx, xbTime, stSumMemo, brs), NowTime, MyBase.CompanyTitle));
                        sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_入库跟催表。", NowTime, MyBase.CompanyTitle));

                        HSSFWorkbook wb = new HSSFWorkbook();
                        string[] fNames = new string[] { "DepartmentName", "ProcessName", "MachineName", "ProduceRdsNo", "AccNoteRdsNo", "ProductName", "Numb1", "RptDate", "InStockTimeSpan", "StockLocationTypeName", "ProducePropertyName" };
                        string[] fTitles = new string[] { "部门", "工序", "机台", "工单编号", "验收入库单", "料号", "数量", "日期", "入库耗时(H)", "入仓别", "工单性质" };
                        string[] fSortFields = new string[] { "FullSortID", "ProduceRdsNo", "AccNoteRdsNo" };
                        ExportToExcelIPQCAndScrapBack(wb, md, fNames, fTitles, "入库跟催表", fSortFields);
                        string tmpFileName = Path.GetTempFileName();
                        MyRecord.Say(string.Format("tmpFileName = {0}", tmpFileName));
                        Regex rgx = new System.Text.RegularExpressions.Regex(@"(?<=tmp)(.+)(?=\.tmp)");
                        string tmpFileNameLast = rgx.Match(tmpFileName).Value;
                        MyRecord.Say(string.Format("tmpFileNameLast = {0}", tmpFileNameLast));
                        string dName = string.Format("{0}\\TMPExcel", Application.StartupPath);
                        if (!Directory.Exists(dName)) Directory.CreateDirectory(dName);
                        string fname = string.Format("{0}\\{1}", dName, string.Format("NOTICE_{0:yyyyMMdd}_tmp{1}.xls", NowTime.Date, tmpFileNameLast));
                        MyRecord.Say(string.Format("fname = {0}", fname));
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

                        wb.Close();
                        MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", md.Rows.Count));


                        //string mailto = ConfigurationManager.AppSettings["InStockFollowUPMailTo"], mailcc = ConfigurationManager.AppSettings["InStockFollowUPMailCC"];
                        //MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        //sm.MailTo = mailto; // "my18@my.imedia.com.tw,xang@my.imedia.com.tw,lghua@my.imedia.com.tw,my64@my.imedia.com.tw";
                        //sm.MailCC = mailcc; // "jane123@my.imedia.com.tw,lwy@my.imedia.com.tw,my80@my.imedia.com.tw";
                        ////sm.MailTo = "my80@my.imedia.com.tw";

                        MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("InStockFollowUP");
                        MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                        sm.MailTo = mAddress.MailTo;
                        sm.MailCC = mAddress.MailCC;
                        //sm.MailTo = "my80@my.imedia.com.tw";

                        MyRecord.Say("发送邮件。");
                        sm.SendOut();
                        sm.mail.Dispose();
                        sm = null;
                        MyRecord.Say("已经发送。");
                        if (File.Exists(fname)) File.Delete(fname);
                    }
                    else
                    {
                        MyRecord.Say("没有找到资料。");
                        MyBase.SendMail sm = new MyBase.SendMail();
                        MyRecord.Say("加载邮件内容。");

                        sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, string.Format(bx2, xbTime), NowTime, MyBase.CompanyTitle));
                        sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_入库跟催表。", NowTime, MyBase.CompanyTitle));

                        //string mailto = ConfigurationManager.AppSettings["InStockFollowUPMailTo"], mailcc = ConfigurationManager.AppSettings["InStockFollowUPMailCC"];
                        //MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        //sm.MailTo = mailto;
                        //sm.MailCC = mailcc;
                        ////sm.MailTo = "my80@my.imedia.com.tw";
                        ////sm.MailCC = "my80@my.imedia.com.tw";

                        MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("InStockFollowUP");
                        MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                        sm.MailTo = mAddress.MailTo;
                        sm.MailCC = mAddress.MailCC;
                        //sm.MailTo = "my80@my.imedia.com.tw";


                        MyRecord.Say("发送邮件。");
                        sm.SendOut();
                        sm.mail.Dispose();
                        sm = null;
                        MyRecord.Say("已经发送。");
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

        #region 发送差异数

        void SendProduceDiffNumbEmailLoder()
        {
            MyRecord.Say("开启定时发差异数进程..........");
            Thread t = new Thread(SendProduceDiffNumbEmail);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启定时差异数进程完成。");
        }

        #region 计算要用到的类

        class ProduceListDataItem
        {
            public ProduceListDataItem(MyData.MyDataRow r)
            {
                RdsNo = r.Value("RdsNo");
                CUSTORDERNO = Convert.ToString(r["CUSTORDERNO"]);
                CustCode = Convert.ToString(r["CustCode"]);
                ProdCode = Convert.ToString(r["ProdCode"]);
                ProdName = Convert.ToString(r["ProdName"]);
                PlateCode = Convert.ToString(r["PlateCode"]);
                OrderNo = Convert.ToString(r["OrderNo"]);
                pDeliveryDate = Convert.ToDateTime(r["pDeliveryDate"]);
                FinishDate = Convert.ToDateTime(r["FinishDate"]);
                StockDate = Convert.ToDateTime(r["StockDate"]);
                PNumb = Convert.ToDouble(r["PNumb"]);
                InStockNumb = Convert.ToDouble(r["InStockNumb"]);
                FinishNumb = Convert.ToDouble(r["FinishNumb"]);
                PropertyName = Convert.ToString(r["PropertyName"]);
                StatusName = Convert.ToString(r["StatusName"]);
                Inputer = Convert.ToString(r["Inputer"]);
                InputDate = Convert.ToDateTime(r["InputDate"]);
                SRemark = Convert.ToString(r["SRemark"]);
                RMan = Convert.ToString(r["RMan"]);
                Status = Convert.ToString(r["Status"]);
                StockStatus = Convert.ToString(r["StockStatus"]);
                pid = Convert.ToString(r["pid"]);
                Property = Convert.ToInt32(r["Property"]);
                id = Convert.ToInt32(r["id"]);
                LastProcessID = Convert.ToInt32(r["LastProcID"]);
                PhoneSubject = Convert.ToString(r["PhoneSubject"]);
                Size = Convert.ToString(r["Size"]);
                TypeName = Convert.ToString(r["TypeName"]);
                Secret = Convert.ToString(r["Secret"]);
                Bonded = Convert.ToString(r["Bonded"]);
                DeclaretionName = Convert.ToString(r["DeclaretionName"]);
                Description = Convert.ToString(r["Description"]);
                isKit = Convert.ToInt32(r["isKit"]) == 1;
                isStopStock = Convert.ToInt32(r["StopStock"]) == 1;
                isFinalized = Convert.ToInt32(r["finalized"]) == 1;
            }
            public string RdsNo { get; set; }
            public string CUSTORDERNO { get; set; }
            public string CustCode { get; set; }
            public string ProdCode { get; set; }
            public string ProdName { get; set; }
            public string PlateCode { get; set; }
            public string OrderNo { get; set; }
            public DateTime pDeliveryDate { get; set; }
            public DateTime FinishDate { get; set; }
            public DateTime StockDate { get; set; }
            public double PNumb { get; set; }
            public double InStockNumb { get; set; }
            public double FinishNumb { get; set; }
            public string PropertyName { get; set; }
            public string StatusName { get; set; }
            public string Inputer { get; set; }
            public DateTime InputDate { get; set; }
            public string SRemark { get; set; }
            public string RMan { get; set; }
            public string Status { get; set; }
            public string StockStatus { get; set; }
            public string pid { get; set; }
            public int Property { get; set; }
            public int id { get; set; }
            public int LastProcessID { get; set; }
            public string PhoneSubject { get; set; }
            public string Size { get; set; }
            public string TypeName { get; set; }
            public string Secret { get; set; }
            public string Bonded { get; set; }
            public string DeclaretionName { get; set; }
            public string Description { get; set; }
            public bool isKit { get; set; }
            public bool isStopStock { get; set; }
            public bool isFinalized { get; set; }
        }

        class ProduceProcessDataItem
        {
            public ProduceProcessDataItem(MyData.MyDataRow r)
            {
                RdsNo = Convert.ToString(r["RdsNo"]);
                id = Convert.ToInt32(r["id"]);
                ProcessCode = Convert.ToString(r["ProcessCode"]);
                ReqNumb = Convert.ToDouble(r["ReqNumb"]);
                FinishNumb = Convert.ToDouble(r["FinishNumb"]);
                isOut = Convert.ToInt32(r["isOut"]) == 1;
                ProcRemark = Convert.ToString(r["ProcRemark"]);
                Remark = Convert.ToString(r["Remark"]);
                MachineCode = Convert.ToString(r["MachineCode"]);
                PType = Convert.ToInt32(r["PType"]);
                PartID = Convert.ToString(r["PartID"]);
                OverDate = Convert.ToDateTime(r["OverDate"]);
                WastageNumb = Convert.ToDouble(r["WastageNumb"]);
                RejectNumb = Convert.ToDouble(r["RejectNumb"]);
                PreviousID = Convert.ToInt32(r["PreviousID"]);
                NextID = Convert.ToInt32(r["NextID"]);
                PrintPartID = Convert.ToString(r["PrintPartID"]);
                LossedNumb = Convert.ToDouble(r["LossedNumb"]);
                ColNumb = Convert.ToInt32(r["ColNumb"]);
                UseStockNumb = Convert.ToDouble(r["UseStockNumb"]);
                MachineName = Convert.ToString(r["MachineName"]);
                DepartmentName = Convert.ToString(r["DepartmentName"]);
                ProcessName = Convert.ToString(r["ProcessName"]);
            }
            public string RdsNo { get; set; }
            public int id { get; set; }
            public string ProcessCode { get; set; }
            public double ReqNumb { get; set; }
            public double FinishNumb { get; set; }
            public bool isOut { get; set; }
            public string ProcRemark { get; set; }
            public string Remark { get; set; }
            public string MachineCode { get; set; }
            public int PType { get; set; }
            public string PartID { get; set; }
            public DateTime OverDate { get; set; }
            public double WastageNumb { get; set; }
            public double RejectNumb { get; set; }
            public int PreviousID { get; set; }
            public int NextID { get; set; }
            public string PrintPartID { get; set; }
            public double LossedNumb { get; set; }
            public int ColNumb { get; set; }
            public double UseStockNumb { get; set; }
            public double PickedNumb { get; set; }

            public string MachineName { get; set; }
            public string DepartmentName { get; set; }
            public string ProcessName { get; set; }
        }

        class ProduceMaterialDateItem
        {
            public ProduceMaterialDateItem(MyData.MyDataRow r)
            {
                RdsNo = Convert.ToString(r["RdsNo"]);
                PartID = Convert.ToString(r["PartID"]);
                id = Convert.ToInt32(r["id"]);
                Code = Convert.ToString(r["Code"]);
                name = Convert.ToString(r["name"]);
                Numb = Convert.ToDouble(r["Numb"]);
                PickedNumb = Convert.ToDouble(r["PickedNumb"]);
                OverflowNumb = Convert.ToDouble(r["OverflowNumb"]);
                ReturnNumb = Convert.ToDouble(r["ReturnNumb"]);
                ColNumb = Convert.ToInt32(r["ColNumb"]);
                CutNumb = Convert.ToInt32(r["CutNumb"]);
                UnitNumb = Convert.ToDouble(r["UnitNumb"]);
                ProcID = Convert.ToInt32(r["ProcID"]);
                ProcCode = Convert.ToString(r["ProcCode"]);
                isPaper = Convert.ToInt32(r["isPaper"]) == 1;
                isInPaper = Convert.ToInt32(r["isInPaper"]) == 1;
                isWavePaper = Convert.ToInt32(r["isWavePaper"]) == 1;

                PickedNumbSum = (PickedNumb + OverflowNumb - ReturnNumb);
                if (isPaper)
                    FinishNumbSum = PickedNumbSum * ColNumb * CutNumb;
                else if (isWavePaper)
                    FinishNumbSum = PickedNumbSum * ColNumb;
                else if (UnitNumb != 0)
                    FinishNumbSum = Math.Round(PickedNumbSum / UnitNumb, 2);
                else
                    FinishNumbSum = PickedNumbSum;
            }
            public string RdsNo { get; set; }
            public int id { get; set; }
            public string Code { get; set; }
            public double Numb { get; set; }
            public double PickedNumb { get; set; }
            public double OverflowNumb { get; set; }
            public double ReturnNumb { get; set; }
            public int ColNumb { get; set; }
            public int CutNumb { get; set; }
            public double UnitNumb { get; set; }
            public int ProcID { get; set; }
            public string ProcCode { get; set; }
            public bool isPaper { get; set; }
            public bool isWavePaper { get; set; }
            public bool isInPaper { get; set; }
            public string PartID { get; set; }
            public string name { get; private set; }
            public double PickedNumbSum { get; private set; }
            public double FinishNumbSum { get; private set; }
        }

        class GridItem
        {
            public string rdsno { get; set; }
            public string CustID { get; set; }
            public string ProdName { get; set; }
            public string ProdCode { get; set; }
            public string partid { get; set; }
            public string DeptmantName { get; set; }
            public string DepartmentCode { get; set; }
            public string procname { get; set; }
            public string machinename { get; set; }
            public double reqnumb { get; set; }
            public double UseStockNumb { get; set; }
            public double picknumb { get; set; }
            public double finishnumb { get; set; }
            public double rejectnumb { get; set; }
            public double wastagenumb { get; set; }
            public double defnumb { get; set; }
            public DateTime overdate { get; set; }
            public string finishword { get; set; }
            public double yield { get; set; }
            public int id { get; set; }
            public int pid { get; set; }
            public string proccode { get; set; }
            public string machinecode { get; set; }
            public string Code { get; set; }
            public DateTime MaxInputTiime { get; set; }
            public DateTime FinishDate { get; set; }
            public double InStockNumb { get; set; }
            public bool Finalized { get; set; }
            public int ColNumb { get; set; }
            public int CutNumb { get; set; }

            public int PropertyID { get; set; }

            public string PropertyName { get; set; }
        }
        #endregion
        void SendProduceDiffNumbEmail()
        {
            try
            {
                MyRecord.Say("-----------------开启定时发送差异数-------------------------");
                #region 邮件体
                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{2}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;</FONT></DIV>
{0}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{1:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");
                string bodyGrid = @"<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	序号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	客户编号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	产品编号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	产品料号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单类别
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	部件
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工序
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	機臺
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	模數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	領料數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	完工數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	不良數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	損耗數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	差異數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	完工
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	良品率
	</TD>
	</TR>
	{0}
</TBODY></TABLE>";
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
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{15}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{16}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{17}
	</TD>
	</TR>
";
                #endregion
                MyRecord.Say("1.设定数据条件，准备开始计算。");
                DateTime NowTime = DateTime.Now;
                DateTime beginTime = NowTime.AddMonths(-6).AddDays(1 - NowTime.Day).Date, endTime = NowTime.AddDays(-10).Date.AddMilliseconds(-10);
                //beginTime = NowTime.AddMonths(-1).AddDays(1 - NowTime.Day).Date;
                DateTime stTime = new DateTime(2017, 1, 1).Date;
                if (CompanyType == "MD")
                {
                    stTime = new DateTime(2017, 4, 1).Date;
                }
                if (beginTime < stTime)
                {
                    beginTime = stTime;
                }
                #region 加载数据源
                Thread.Sleep(200);
                MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
					{
						new MyData.MyDataParameter("@BTime",beginTime, MyData.MyDataParameter.MyDataType.DateTime),
						new MyData.MyDataParameter("@ETime",endTime, MyData.MyDataParameter.MyDataType.DateTime),
						new MyData.MyDataParameter("@ProduceRdsNo",string.Empty),
						new MyData.MyDataParameter("@PName",string.Empty),
						new MyData.MyDataParameter("@PCode",string.Empty),
						new MyData.MyDataParameter("@CUST",string.Empty),
						new MyData.MyDataParameter("@FinishStatus",0, MyData.MyDataParameter.MyDataType.Int),
						new MyData.MyDataParameter("@StockStatus",3, MyData.MyDataParameter.MyDataType.Int),
						new MyData.MyDataParameter("@ProdType",string.Empty),
						new MyData.MyDataParameter("@PhoneSubject",string.Empty),
						new MyData.MyDataParameter("@BK",-1, MyData.MyDataParameter.MyDataType.Int),
						new MyData.MyDataParameter("@Bond",-1, MyData.MyDataParameter.MyDataType.Int)
					};
                MyRecord.Say(string.Format("1.数据条件：beginTime={0},endTime={1}", beginTime, endTime));
                MyRecord.Say("1.从后台读取数据。");
                MyRecord.Say("第一步，按设定条件获取工单....");
                DateTime bTime = DateTime.Now;
                string SQL = @"EXEC [_PMC_Progress_List] @BTime,@ETime,@ProduceRdsNo,@PCode,@PName,@CUST,@ProdType,@FinishStatus,@StockStatus,@PhoneSubject,@Bond,@BK";
                MyData.MyDataTable md = new MyData.MyDataTable(SQL, mps);
                MyRecord.Say(string.Format("第一步，工单获取耗时：{0}分钟。", (DateTime.Now - bTime).TotalMinutes));
                MyRecord.Say("第二步，按设定条件获取工单，工艺信息和完工数量....");
                Thread.Sleep(200);
                bTime = DateTime.Now;
                SQL = @"EXEC [_PMC_Progress_List_Process] @BTime,@ETime,@ProduceRdsNo,@PCode,@PName,@CUST,@ProdType,@FinishStatus,@StockStatus,@PhoneSubject,@Bond,@BK";
                MyData.MyDataTable md2 = new MyData.MyDataTable(SQL, mps);
                MyRecord.Say(string.Format("第二步，工单工艺获取耗时：{0}分钟。", (DateTime.Now - bTime).TotalMinutes));
                MyRecord.Say("第三步，按设定条件获取工单，原辅料信息和领用数量....");
                Thread.Sleep(200);
                bTime = DateTime.Now;
                SQL = @"EXEC [_PMC_Progress_List_Material] @BTime,@ETime,@ProduceRdsNo,@PCode,@PName,@CUST,@ProdType,@FinishStatus,@StockStatus,@PhoneSubject,@Bond,@BK";
                MyData.MyDataTable md3 = new MyData.MyDataTable(SQL, mps);
                MyRecord.Say(string.Format("第三步，工单物料获取耗时：{0}分钟。", (DateTime.Now - bTime).TotalMinutes));
                Thread.Sleep(200);

                #endregion
                if (md != null && md.MyRows.Count > 0)
                {
                    #region 保存到内存

                    List<ProduceListDataItem> ProduceListDataSource = new List<ProduceListDataItem>();
                    List<ProduceProcessDataItem> ProduceProcessDataSource = new List<ProduceProcessDataItem>();
                    List<ProduceMaterialDateItem> ProduceMaterialDataSource = new List<ProduceMaterialDateItem>();
                    List<GridItem> GridDataSource = new List<GridItem>();

                    MyRecord.Say("第四步，数据已经读取，正在保存到本地....");
                    Thread.Sleep(200);
                    var vMain = from a in md.MyRows
                                orderby Convert.ToInt32(a["ID"]) // Convert.ToString(a["ProduceRdsNo"])
                                select new ProduceListDataItem(a);
                    ProduceListDataSource = vMain.ToList();

                    if (md2 != null && md2.MyRows.Count > 0)
                    {
                        var vProcess = from a in md2.MyRows
                                       select new ProduceProcessDataItem(a);
                        ProduceProcessDataSource = vProcess.ToList();
                    }

                    if (md3 != null && md3.MyRows.Count > 0)
                    {
                        var vMaterial = from a in md3.MyRows
                                        select new ProduceMaterialDateItem(a);
                        ProduceMaterialDataSource = vMaterial.ToList();
                    }
                    MyRecord.Say(string.Format("第四步，保存到本地耗时：{0}分钟。", (DateTime.Now - bTime).TotalMinutes));
                    #endregion
                    #region 计算

                    GridDataSource.Clear();
                    List<GridItem> OneGridData = new List<GridItem>();
                    double xNoteIndex = 1, xNoteCount = ProduceListDataSource.Count();
                    MyRecord.Say(string.Format("2.在本机计算表格内容。共{0}行。", xNoteCount));
                    foreach (var xNote in ProduceListDataSource)
                    {
                        OneGridData.Clear();
                        bTime = DateTime.Now;
                        var vxPart = from a in ProduceProcessDataSource
                                     where a.RdsNo.ToUpper() == xNote.RdsNo.ToUpper()
                                     group a by a.PartID into g
                                     orderby g.Key == "" ? 1 : 0, g.Key
                                     select g.Key;
                        foreach (var xPart in vxPart)
                        {
                            ///开始加载工艺
                            var vxp = from a in ProduceProcessDataSource
                                      where a.PartID == xPart && a.RdsNo == xNote.RdsNo
                                      orderby a.id
                                      select a;
                            if (vxp != null && vxp.Count() > 0)
                            {
                                foreach (var vPi in vxp)
                                {
                                    GridItem gp = new GridItem();
                                    double pickedNumber = 0;
                                    gp.picknumb = 0;
                                    if (vPi.PreviousID > 0)
                                    {
                                        if (vPi.PartID.Length <= 0)
                                        {
                                            var vpick = from a in ProduceProcessDataSource
                                                        where a.RdsNo.ToUpper() == xNote.RdsNo.ToUpper() && a.NextID == vPi.id
                                                        select a.FinishNumb + a.UseStockNumb;
                                            if (vpick != null && vpick.Count() > 0) pickedNumber = vpick.Min();
                                        }
                                        else
                                        {
                                            var vpick = from a in ProduceProcessDataSource
                                                        where a.RdsNo.ToUpper() == xNote.RdsNo.ToUpper() && a.id == vPi.PreviousID
                                                        select a.FinishNumb + a.UseStockNumb;
                                            if (vpick != null && vpick.Count() > 0) pickedNumber = vpick.Min();
                                        }
                                    }
                                    else
                                    {
                                        if (vPi.ProcessCode == "9050")
                                        {
                                            if (xNote.InStockNumb == 0 && vPi.FinishNumb == 0 && xNote.FinishDate > Convert.ToDateTime("2000-01-01"))
                                                pickedNumber = 0;
                                            else
                                            {
                                                if (xNote.FinishDate > Convert.ToDateTime("2000-01-01"))
                                                    pickedNumber = Math.Floor(vPi.FinishNumb) + Math.Floor(vPi.RejectNumb) + Math.Floor(vPi.LossedNumb);
                                                else
                                                    pickedNumber = vPi.ReqNumb;
                                            }
                                        }
                                        else
                                        {
                                            pickedNumber = (vPi.ReqNumb + vPi.WastageNumb) * vPi.ColNumb;
                                            if (xNote.InputDate > Convert.ToDateTime("2014-11-18").Date)
                                            {
                                                var vxm = from a in ProduceMaterialDataSource
                                                          where a.isPaper && a.PartID == xPart && a.RdsNo == xNote.RdsNo && a.ProcID == vPi.id
                                                          orderby a.id
                                                          select a;
                                                //纸张
                                                if (vxm != null && vxm.Count() > 0)
                                                {
                                                    foreach (var xMi in vxm)
                                                    {
                                                        GridItem g = new GridItem();
                                                        g.rdsno = xMi.RdsNo;
                                                        g.CustID = xNote.CustCode;
                                                        g.ProdName = xNote.ProdName;
                                                        g.ProdCode = xNote.ProdCode;
                                                        g.partid = xPart.Length > 0 ? xPart : ".";
                                                        g.DeptmantName = MyConvert.ZHLC("领料");
                                                        g.procname = vPi.ProcessName; //MyServiceLogic.Processes.Contains(xMi.ProcCode) ? MyServiceLogic.Processes[xMi.ProcCode].Name : string.Empty;
                                                        g.machinename = xMi.name;
                                                        g.reqnumb = xMi.Numb;
                                                        g.UseStockNumb = 0;
                                                        g.picknumb = xMi.PickedNumbSum;
                                                        g.finishnumb = xMi.FinishNumbSum;
                                                        g.rejectnumb = 0;
                                                        g.wastagenumb = 0;
                                                        g.defnumb = 0;
                                                        g.overdate = DateTime.MinValue;
                                                        g.finishword = null;
                                                        g.yield = 0;
                                                        g.id = 0;
                                                        g.pid = 0;
                                                        g.proccode = xMi.ProcCode;
                                                        g.machinecode = null;
                                                        g.Code = xNote.ProdCode;
                                                        g.MaxInputTiime = DateTime.MinValue;
                                                        g.FinishDate = xNote.StockDate;
                                                        g.InStockNumb = xNote.InStockNumb;
                                                        g.Finalized = xNote.isFinalized;
                                                        g.ColNumb = xMi.ColNumb;
                                                        g.CutNumb = xMi.CutNumb;
                                                        GridDataSource.Add(g);
                                                        OneGridData.Add(g);
                                                    }
                                                    pickedNumber = vxm.Sum(u => u.FinishNumbSum);
                                                }
                                                else
                                                {
                                                    if (xNote.InStockNumb == 0 && vPi.FinishNumb == 0 && xNote.FinishDate > Convert.ToDateTime("2000-01-01"))
                                                        pickedNumber = 0;
                                                    else
                                                        pickedNumber = Math.Floor(vPi.FinishNumb) + Math.Floor(vPi.RejectNumb) + Math.Floor(vPi.LossedNumb);
                                                    #region 以后再说。
                                                    ////其他
                                                    //var vxmb = from a in ProduceMaterialDataSource
                                                    //           where a.PartID == xPart && a.RdsNo.ToUpper() == xNote.RdsNo.ToUpper() && a.ProcID == vPi.id
                                                    //           orderby a.FinishNumbSum, a.id
                                                    //           select a;
                                                    //if (vxmb != null && vxmb.Count() > 0)
                                                    //{
                                                    //    var xMi = vxmb.FirstOrDefault();
                                                    //    GridItem g = new GridItem();
                                                    //    g.rdsno = xMi.RdsNo;
                                                    //    g.CustID = xNote.CustCode;
                                                    //    g.ProdName = xNote.ProdName;
                                                    //    g.ProdCode = xNote.ProdCode;
                                                    //    g.partid = xPart.Length > 0 ? xPart : ".";
                                                    //    g.DeptmantName = LCStr("领料");
                                                    //    g.procname = MyServiceLogic.Processes.Contains(xMi.ProcCode) ? MyServiceLogic.Processes[xMi.ProcCode].Name : string.Empty;
                                                    //    g.machinename = xMi.name;
                                                    //    g.reqnumb = xMi.Numb;
                                                    //    g.UseStockNumb = 0;
                                                    //    g.picknumb = xMi.PickedNumbSum;
                                                    //    g.finishnumb = xMi.FinishNumbSum;
                                                    //    g.rejectnumb = 0;
                                                    //    g.wastagenumb = 0;
                                                    //    g.defnumb = 0;
                                                    //    g.overdate = DateTime.MinValue;
                                                    //    g.finishword = null;
                                                    //    g.yield = 0;
                                                    //    g.id = 0;
                                                    //    g.pid = 0;
                                                    //    g.proccode = xMi.ProcCode;
                                                    //    g.machinecode = null;
                                                    //    g.Code = xNote.ProdCode;
                                                    //    g.MaxInputTiime = DateTime.MinValue;
                                                    //    g.FinishDate = xNote.StockDate;
                                                    //    GridDataSource.Add(g);
                                                    //    pickedNumber = xMi.FinishNumbSum;
                                                    //}
                                                    #endregion
                                                }
                                            }
                                            else
                                            {
                                                if (xNote.InStockNumb == 0 && vPi.FinishNumb == 0 && xNote.FinishDate > Convert.ToDateTime("2000-01-01"))
                                                    pickedNumber = 0;
                                                else
                                                    pickedNumber = Math.Floor(vPi.FinishNumb) + Math.Floor(vPi.RejectNumb) + Math.Floor(vPi.LossedNumb);
                                            }
                                        }
                                    }
                                    gp.picknumb = vPi.PickedNumb = pickedNumber;
                                    gp.rdsno = vPi.RdsNo;
                                    gp.CustID = xNote.CustCode;
                                    gp.ProdName = xNote.ProdName;
                                    gp.ProdCode = xNote.ProdCode;
                                    gp.partid = xPart.Length > 0 ? xPart : ".";
                                    gp.DeptmantName = vPi.DepartmentName;
                                    gp.machinename = vPi.MachineName;
                                    gp.procname = vPi.ProcessName;
                                    gp.reqnumb = vPi.ReqNumb;
                                    gp.UseStockNumb = vPi.UseStockNumb;
                                    gp.finishnumb = vPi.FinishNumb;
                                    gp.rejectnumb = vPi.RejectNumb;
                                    gp.wastagenumb = vPi.LossedNumb;
                                    double defNumb = Math.Floor(pickedNumber) - (Math.Floor(vPi.FinishNumb) + Math.Floor(vPi.RejectNumb) + Math.Floor(vPi.LossedNumb));
                                    gp.defnumb = defNumb;
                                    gp.overdate = vPi.OverDate;
                                    gp.finishword = vPi.OverDate > DateTime.MinValue ? "已完" : "";
                                    if (pickedNumber > 0)
                                        gp.yield = vPi.FinishNumb / pickedNumber;
                                    else
                                        gp.yield = 0;
                                    gp.id = xNote.id;
                                    gp.pid = vPi.id;
                                    gp.proccode = vPi.ProcessCode;
                                    gp.machinecode = vPi.MachineCode;
                                    gp.Code = xNote.ProdCode;
                                    gp.MaxInputTiime = DateTime.MinValue;
                                    gp.FinishDate = xNote.StockDate;
                                    gp.InStockNumb = xNote.InStockNumb;
                                    gp.Finalized = xNote.isFinalized;
                                    gp.ColNumb = vPi.ColNumb;
                                    gp.PropertyID = xNote.Property;
                                    gp.PropertyName = xNote.PropertyName;
                                    GridDataSource.Add(gp);
                                    OneGridData.Add(gp);
                                }
                            }
                        }

                        #region 加载入库行
                        if (!xNote.isStopStock)
                        {
                            GridItem lg = new GridItem();
                            lg.rdsno = xNote.RdsNo;
                            lg.CustID = xNote.CustCode;
                            lg.ProdName = xNote.ProdName;
                            lg.ProdCode = xNote.ProdCode;
                            lg.partid = ".";
                            lg.DeptmantName = null;
                            lg.procname = MyConvert.ZHLC("成品入库");
                            lg.machinename = null;
                            lg.reqnumb = xNote.PNumb;
                            lg.UseStockNumb = 0;
                            double l_pickedNumber = 0, l_firstPickedNumber = 0;
                            if (ProduceProcessDataSource != null)
                            {
                                var vl = from a in ProduceProcessDataSource
                                         where a.RdsNo.ToUpper() == xNote.RdsNo.ToUpper() && a.NextID == 0 && a.id == xNote.LastProcessID
                                         select a.FinishNumb; //+ a.UseStockNumb;
                                if (vl != null && vl.Count() > 0)
                                {
                                    l_pickedNumber = vl.FirstOrDefault();
                                    lg.picknumb = l_pickedNumber;
                                }
                                else
                                    lg.picknumb = 0;
                                var vfl = from a in ProduceProcessDataSource
                                          where a.RdsNo == xNote.RdsNo && a.PreviousID == 0
                                          select a.PickedNumb;
                                if (vfl != null && vfl.Count() > 0)
                                {
                                    l_firstPickedNumber = vfl.Min();
                                }
                            }
                            else
                            {
                                lg.picknumb = 0;
                            }
                            lg.finishnumb = xNote.InStockNumb;
                            lg.rejectnumb = 0;
                            lg.wastagenumb = 0;
                            double l_defNumb = Math.Floor(l_pickedNumber - xNote.InStockNumb);
                            lg.defnumb = l_defNumb;
                            lg.overdate = xNote.StockDate;
                            lg.finishword = xNote.StatusName;
                            if (l_firstPickedNumber > 0)
                                lg.yield = xNote.InStockNumb / l_firstPickedNumber;
                            else
                                lg.yield = 0;
                            lg.id = xNote.id;
                            lg.pid = 0;
                            lg.proccode = null;
                            lg.machinecode = null;
                            lg.Code = xNote.ProdCode;
                            lg.MaxInputTiime = DateTime.MinValue;
                            lg.FinishDate = xNote.StockDate;
                            lg.InStockNumb = xNote.InStockNumb;
                            lg.Finalized = xNote.isFinalized;
                            lg.ColNumb = 1;
                            lg.PropertyID = xNote.Property;
                            lg.PropertyName = xNote.PropertyName;
                            GridDataSource.Add(lg);
                            OneGridData.Add(lg);
                        }
                        #endregion

                        var nvk1 = from a in OneGridData
                                   where a.defnumb != 0 && a.FinishDate > Convert.ToDateTime("2000-01-01") && a.InStockNumb > 0 && !a.Finalized
                                   select a;
                        string fWord = string.Empty;
                        if (nvk1 != null && nvk1.Count() > 0)
                        {
                            GridItem gi = nvk1.FirstOrDefault();
                            if (gi.Finalized)
                            {
                                MyConvert.ZHLC(string.Format("有差异，被关闭，工序：{0}，机台：{1}，差异数：{2}", gi.procname, gi.machinename, gi.defnumb));
                            }
                            else if (gi.PropertyID == 5)
                            {
                                MyConvert.ZHLC(string.Format("有差异，忽略{3}，工序：{0}，机台：{1}，差异数：{2}", gi.procname, gi.machinename, gi.defnumb, gi.PropertyName));
                            }
                            else
                            {
                                fWord = MyConvert.ZHLC(string.Format("有差异 工序：{0}，机台：{1}，差异数：{2}", gi.procname, gi.machinename, gi.defnumb));
                            }
                        }
                        else
                        {
                            fWord = MyConvert.ZHLC("无差异");
                        }
                        double xPercent = xNoteIndex / xNoteCount;
                        MyRecord.Say(string.Format("第{3:0}条，工单号：{0}已经处理，耗时：{1:0.0}秒，完成{4:0.00%}，{2}", xNote.RdsNo, (DateTime.Now - bTime).TotalSeconds, fWord, xNoteIndex, xPercent));
                        xNoteIndex++;
                    }

                    #endregion
                    #region 筛选加工
                    MyRecord.Say("对数据进行筛选加工。");
                    List<GridItem> _ThisGridDataSource = new List<GridItem>();
                    DateTime tmDateTime = new DateTime(1998, 1, 1);
                    _ThisGridDataSource = GridDataSource;
                    var nvk = from a in _ThisGridDataSource
                              where a.defnumb != 0 && a.FinishDate > tmDateTime && a.InStockNumb > 0 && !a.Finalized && a.PropertyID != 5
                              group a by a.rdsno into g
                              select g.Key;
                    string[] mm = nvk.ToArray();
                    var nvv = from b in _ThisGridDataSource
                              where mm.Contains(b.rdsno)
                              orderby b.rdsno
                              select b;
                    _ThisGridDataSource = nvv.ToList();

                    #endregion
                    MyRecord.Say("3.生成邮件内容。");
                    int EmailGridRow = 0;
                    string brs = "";
                    try
                    {
                        foreach (var item in _ThisGridDataSource)
                        {
                            if (Math.Floor(item.defnumb) != 0)
                            {
                                brs += string.Format(br, EmailGridRow + 1,
                                                     item.rdsno,
                                                     item.CustID,
                                                     item.Code,
                                                     item.ProdName,
                                                     item.PropertyName,
                                                     item.partid,
                                                     item.DeptmantName,
                                                     item.procname,
                                                     item.machinename,
                                                     item.ColNumb,
                                                     item.picknumb,
                                                     item.finishnumb,
                                                     item.rejectnumb,
                                                     item.wastagenumb,
                                                     item.defnumb,
                                                     item.finishword,
                                                     string.Format("{0:0.00%}", item.yield)
                                                     );
                                EmailGridRow++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MyRecord.Say(ex);
                    }
                    MyRecord.Say("3.生成邮件内容完成。");
                    MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", EmailGridRow));

                    MyRecord.Say("创建SendMail。");
                    MyBase.SendMail sm = new MyBase.SendMail();
                    MyRecord.Say("加载邮件内容。");
                    string xBodyString = string.Empty;
                    if (EmailGridRow > 0)
                    {
                        xBodyString = string.Format(
@"<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容为{0:yy/MM/dd HH:mm}至{1:yy/MM/dd HH:mm}之间完工有差异的工单。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (为了减小邮件大小，表格中只包含有差异数的工序。详细情况请在ERP生产进度中查询。)</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
{2}</FONT>
</DIV>", beginTime, endTime, string.Format(bodyGrid, brs));
                    }
                    else
                    {
                        xBodyString = string.Format(@"<DIV><FONT size=4 color=#FF860900 face=PMingLiU >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {0:yy/MM/dd HH:mm}至{1:yy/MM/dd HH:mm}之间没有发现结单差异之工单。</FONT><Br><FONT size=4 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </FONT><FONT size=5 color=#FF860900 face=PMingLiU>请再接再厉！</FONT></DIV>", beginTime, endTime);
                    }
                    sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, xBodyString, NowTime, MyBase.CompanyTitle));
                    sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}結單差異數提醒。", NowTime, MyBase.CompanyTitle));
                    //string mailto = ConfigurationManager.AppSettings["FinishMailTo"], mailcc = ConfigurationManager.AppSettings["FinishMailCC"];
                    //MyRecord.Say(string.Format("发送邮件地址：\r\n MailTO:{0} \r\n MailCC:{1}", mailto, mailcc));
                    //sm.MailTo = mailto;
                    //sm.MailCC = mailcc;
                    ////sm.MailTo = "my80@my.imedia.com.tw";

                    MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("Finish");
                    MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                    sm.MailTo = mAddress.MailTo;
                    sm.MailCC = mAddress.MailCC;
                    //sm.MailTo = "my80@my.imedia.com.tw";

                    MyRecord.Say("发送邮件。");
                    NowTime = DateTime.Now;
                    sm.SendOut();
                    MyRecord.Say(string.Format("发送完成，耗時：{0:#,##0.00}", (DateTime.Now - NowTime).TotalSeconds));

                    MyRecord.Say("------------------发送完成----------------------------");
                }
            }
            catch (Exception e)
            {
                MyRecord.Say(e);
            }
        }


        #region 计算用到的类2

        class GridItem2
        {
            public GridItem2(MyData.MyDataRow r)
            {
                rdsno = r.Value("rdsno");
                CustCode = r.Value("CustCode");
                ProdCode = r.Value("ProdCode");
                ProdName = r.Value("ProdName");
                PhoneSubject = r.Value("PhoneSubject");
                status = r.IntValue("status");
                StockStatus = r.IntValue("StockStatus");
                Bonded = r.IntValue("Bonded");
                Secret = r.IntValue("Secret");
                OrderType = r.Value("OrderType");
                FinishStatusName = r.Value("FinishStatusName");
                StockStatusName = r.Value("StockStatusName");
                OrderTypeName = r.Value("OrderTypeName");
                PropertyName = r.Value("PropertyName");
                PID = r.IntValue("PID");
                PartID = r.Value("PartID");
                MachineName = r.Value("MachineName");
                ProcessName = r.Value("ProcessName");
                DepartmentName = r.Value("DepartmentName");
                reqnumb = r.Value<double>("reqnumb");
                UseStockNumb = r.Value<double>("UseStockNumb");
                ColNumb = r.Value<double>("ColNumb");
                CutNumb = r.Value<double>("CutNumb");
                PickingNumb = r.Value<double>("PickingNumb");
                FinishNumb = r.Value<double>("FinishNumb");
                RejectNumb = r.Value<double>("RejectNumb");
                AdjustNumb = r.Value<double>("AdjustNumb");
                SampleNumb = r.Value<double>("SampleNumb");
                DiffNumb = r.Value<double>("DiffNumb");
                SortID = r.Value<double>("SortID");
                FinishName = r.Value("FinishName");
                InputDate = r.DateTimeValue("InputDate");
                overdate = r.DateTimeValue("overdate");
                PaperCode = r.Value("PaperCode");
                yield = r.Value<double>("yield");
                MaxInputTime = r.DateTimeValue("MaxInputTime");
            }

            public string rdsno { get; set; }
            public string CustCode { get; set; }
            public string ProdCode { get; set; }
            public string ProdName { get; set; }
            public string PhoneSubject { get; set; }
            public int status { get; set; }
            public int StockStatus { get; set; }
            public int Bonded { get; set; }
            public int Secret { get; set; }
            public string OrderType { get; set; }
            public string FinishStatusName { get; set; }
            public string StockStatusName { get; set; }

            public string StatusName
            {
                get
                {
                    return status == -1 ? FinishStatusName : StockStatusName;
                }
            }
            public string OrderTypeName { get; set; }
            public string PropertyName { get; set; }
            public int PID { get; set; }
            public string PartID { get; set; }
            public string MachineName { get; set; }
            public string ProcessName { get; set; }
            public string DepartmentName { get; set; }
            public double reqnumb { get; set; }
            public double UseStockNumb { get; set; }
            public double ColNumb { get; set; }
            public double CutNumb { get; set; }
            public double PickingNumb { get; set; }
            public double FinishNumb { get; set; }
            public double RejectNumb { get; set; }
            public double AdjustNumb { get; set; }
            public double SampleNumb { get; set; }
            public double DiffNumb { get; set; }
            public double SortID { get; set; }
            public string FinishName { get; set; }
            public DateTime InputDate { get; set; }

            public DateTime overdate { get; set; }

            public string PaperCode { get; set; }

            public double yield { get; set; }

            public DateTime MaxInputTime { get; set; }

        }

        #endregion

        void SendProduceDiffNumbEmailLoder2()
        {
            MyRecord.Say("开启定时发差异数进程..........");
            Thread t = new Thread(SendProduceDiffNumbEmail2);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启定时差异数进程完成。");
        }

        void SendProduceDiffNumbEmail2()
        {
            try
            {
                MyRecord.Say("-----------------开启定时发送差异数-------------------------");
                #region 邮件体
                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{2}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;</FONT></DIV>
{0}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{1:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");
                string bodyGrid = @"<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	序号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	客户
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	产品编号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	产品料号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	下单类别
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	部件
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工序
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	機臺
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	模數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	領料數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	完工數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	不良數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	过版纸
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	样品数
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	差異數
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单状态
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	良品率
	</TD>
	</TR>
	{0}
</TBODY></TABLE>";
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
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{15}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{16}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{17}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{18}
	</TD>
	</TR>
";
                #endregion
                MyRecord.Say("1.设定数据条件，准备开始计算。");
                DateTime NowTime = DateTime.Now;
                #region 加载数据源
                Thread.Sleep(200);
                MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
				{
					new MyData.MyDataParameter("@DateBegin","2000-01-01", MyData.MyDataParameter.MyDataType.DateTime),
					new MyData.MyDataParameter("@DateEnd","2020-01-01", MyData.MyDataParameter.MyDataType.DateTime),
					new MyData.MyDataParameter("@RdsNo",null),
					new MyData.MyDataParameter("@Code",null),
					new MyData.MyDataParameter("@Name",null),
					new MyData.MyDataParameter("@CustCode",null),
					new MyData.MyDataParameter("@ProdType",null),
					new MyData.MyDataParameter("@FinishStatus",null),
					new MyData.MyDataParameter("@DepartmentID",0, MyData.MyDataParameter.MyDataType.Int),
					new MyData.MyDataParameter("@ProcessCode",null),
					new MyData.MyDataParameter("@PhoneSubject",null),
					new MyData.MyDataParameter("@BondStatus",-1, MyData.MyDataParameter.MyDataType.Int),
					new MyData.MyDataParameter("@SecretLevel","_ALL"),
					new MyData.MyDataParameter("@StockStatus",3, MyData.MyDataParameter.MyDataType.Int),
					new MyData.MyDataParameter("@DiffFlag",2, MyData.MyDataParameter.MyDataType.Int)
				};
                DateTime bTime = DateTime.Now;
                string SQL = @"EXECUTE [_PMC_Progress_List_New] @DateBegin, @DateEnd, @RdsNo, @Code, @Name, @CustCode, @ProdType, @FinishStatus, @DiffFlag, @DepartmentID, @ProcessCode,@StockStatus, @PhoneSubject, @BondStatus, @SecretLevel";
                MyData.MyDataTable md = new MyData.MyDataTable(SQL, mps);

                #endregion
                if (md != null && md.MyRows.Count > 0)
                {
                    #region 保存到内存

                    var v = from a in md.MyRows select new GridItem2(a);
                    List<GridItem2> GridDataSource2 = v.ToList();

                    var vx = from a in GridDataSource2
                             where a.DiffNumb != 0
                             select a;
                    List<GridItem2> _xGrid = vx.ToList();

                    MyRecord.Say(string.Format("2.读取保存到本地耗时：{0}分钟。", (DateTime.Now - bTime).TotalMinutes));
                    #endregion

                    MyRecord.Say("3.生成邮件内容。");
                    int EmailGridRow = 0;
                    string brs = "";
                    try
                    {
                        foreach (var item in _xGrid)
                        {
                            if (Math.Floor(item.DiffNumb) != 0)
                            {
                                brs += string.Format(br, EmailGridRow + 1,
                                                     item.rdsno,
                                                     item.CustCode,
                                                     item.ProdCode,
                                                     item.ProdName,
                                                     item.OrderTypeName,
                                                     item.PartID,
                                                     item.DepartmentName,
                                                     item.ProcessName,
                                                     item.MachineName,
                                                     item.ColNumb,
                                                     item.PickingNumb,
                                                     item.FinishNumb,
                                                     item.RejectNumb,
                                                     item.AdjustNumb,
                                                     item.SampleNumb,
                                                     item.DiffNumb,
                                                     item.StatusName,
                                                     string.Format("{0:0.00%}", item.yield)
                                                     );
                                EmailGridRow++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MyRecord.Say(ex);
                    }
                    MyRecord.Say("3.生成邮件内容完成。");
                    MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", EmailGridRow));

                    MyRecord.Say("创建SendMail。");
                    MyBase.SendMail sm = new MyBase.SendMail();
                    MyRecord.Say("加载邮件内容。");
                    string xBodyString = string.Empty;
                    if (EmailGridRow > 0)
                    {
                        xBodyString = string.Format(
@"<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容为工单已结（或核销）且有差异的工单。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (下表中仅包含有差异数的工序，详细情况在附档中请参考。)</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
{0}</FONT>
</DIV>", string.Format(bodyGrid, brs));
                    }
                    else
                    {
                        xBodyString = string.Format(@"<DIV><FONT size=4 color=#FF860900 face=PMingLiU >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 没有发现结单差异之工单。</FONT><Br><FONT size=4 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </FONT><FONT size=5 color=#FF860900 face=PMingLiU>各单位再接再励！加油！</FONT></DIV>");
                    }
                    sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, xBodyString, NowTime, MyBase.CompanyTitle));
                    sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}結單差異數提醒。", NowTime, MyBase.CompanyTitle));

                    HSSFWorkbook wb = new HSSFWorkbook();
                    string[] fNames = new string[] { "rdsno", "CustCode", "ProdCode", "ProdName", "OrderTypeName", "PartID", "DepartmentName", "ProcessName", "MachineName", "ColNumb", "reqnumb", "PickingNumb", "FinishNumb", "RejectNumb", "AdjustNumb", "SampleNumb", "DiffNumb", "FinishStatusName", "StockStatusName" };
                    string[] fTitles = new string[] { "工单号", "客户", "产品编号", "产品料号", "下单类型", "部件", "部门", "工序", "机台", "模数", "应产数", "领料数", "完工数", "不良数", "过版纸", "样品数", "差异数", "工单状态", "入库状态" };

                    ExportToExcel(wb, md, fNames, fTitles, "接单有差异工单");

                    string tmpFileName = Path.GetTempFileName();
                    MyRecord.Say(string.Format("tmpFileName = {0}", tmpFileName));
                    Regex rgx = new System.Text.RegularExpressions.Regex(@"(?<=tmp)(.+)(?=\.tmp)");
                    string tmpFileNameLast = rgx.Match(tmpFileName).Value;
                    MyRecord.Say(string.Format("tmpFileNameLast = {0}", tmpFileNameLast));
                    string dName = string.Format("{0}\\TMPExcel", Application.StartupPath);
                    if (!Directory.Exists(dName)) Directory.CreateDirectory(dName);
                    string fname = string.Format("{0}\\{1}", dName, string.Format("NOTICE_{0:yyyyMMdd}_tmp{1}.xls", NowTime.Date, tmpFileNameLast));
                    MyRecord.Say(string.Format("fname = {0}", fname));
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
                    wb.Close();

                    //string mailto = ConfigurationManager.AppSettings["FinishMailTo"], mailcc = ConfigurationManager.AppSettings["FinishMailCC"];
                    //MyRecord.Say(string.Format("发送邮件地址：\r\n MailTO:{0} \r\n MailCC:{1}", mailto, mailcc));
                    //sm.MailTo = mailto;
                    //sm.MailCC = mailcc;
                    //sm.MailTo = "my80@my.imedia.com.tw";

                    MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("Finish");
                    MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                    sm.MailTo = mAddress.MailTo;
                    sm.MailCC = mAddress.MailCC;
                    //sm.MailTo = "my80@my.imedia.com.tw";

                    MyRecord.Say("发送邮件。");
                    NowTime = DateTime.Now;
                    sm.SendOut();
                    sm.mail.Dispose();
                    sm = null;
                    MyRecord.Say("已经发送。");
                    if (File.Exists(fname)) File.Delete(fname);
                    MyRecord.Say(string.Format("发送完成，耗時：{0:#,##0.00}", (DateTime.Now - NowTime).TotalSeconds));
                    MyRecord.Say("------------------发送完成----------------------------");
                }
            }
            catch (Exception e)
            {
                MyRecord.Say(e);
            }
        }




        #endregion

        #region 定时计算完工数
        bool ProduceFeedBackRuning = false;

        void ProduceFeedBackLoder()
        {
            MyRecord.Say("定时计算完工数和领料数线程创建.......");
            Thread t = new Thread(ProduceFeedbackRecord);
            t.IsBackground = true;
            MyRecord.Say("定时计算完工数和领料数线程已经启动。");
            t.Start();
        }

        void ProduceFeedbackRecord()
        {
            try
            {
                ProduceFeedBackRuning = true;
                string SQL = "";
                MyRecord.Say("---------------------启动定时计算完工数。------------------------------");
                DateTime NowTime = DateTime.Now;
                DateTime StartTime = ProduceFeedBackLastRunTime, StopTime = NowTime;
                ProduceFeedBackLastRunTime = NowTime;
                MyRecord.Say(string.Format("计算起始时间：{0:yy/MM/dd HH:mm}，结束时间：{1:yy/MM/dd HH:mm}", StartTime, StopTime));
                MyRecord.Say("定时计算——获取计算范围");
                SQL = @"
							Select Distinct RdsNo From (
							Select Distinct ProduceNo as RdsNo from ProdDailyReport Where InputDate Between @InputBegin And @InputEnd And ProduceNo Like 'PO%'
							Union
							Select Distinct ProductNo as RdsNo from stOutProdlst Where InputDate Between @InputBegin And @InputEnd And ProductNo Like 'PO%'
							Union
							Select Distinct ProduceRdsNo as RdsNo from stOtherOutlst Where CreateDate Between @InputBegin And @InputEnd And ProduceRdsNo Like 'PO%'
							) T
							Where RdsNo is Not Null
							Order by RdsNo
						";
                MyData.MyDataTable mTableProduceFinished = new MyData.MyDataTable(SQL,
                    new MyData.MyDataParameter("@InputBegin", StartTime, MyData.MyDataParameter.MyDataType.DateTime),
                    new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime)
                    );
                if (_StopAll) return;
                if (mTableProduceFinished != null && mTableProduceFinished.MyRows.Count > 0)
                {
                    string memo = string.Format("读取了{0}条记录，下面开始计算....", mTableProduceFinished.MyRows.Count);
                    MyRecord.Say(memo);
                    int mTableProduceFinishedCount = 1;
                    var v = from a in mTableProduceFinished.MyRows
                            orderby Convert.ToString(a["RdsNo"]) descending
                            select a;
                    DateTime mStartTime1 = DateTime.Now;
                    MyData.MyCommand mc = new MyData.MyCommand();
                    foreach (MyData.MyDataRow r in v)
                    {
                        if (_StopAll) return;
                        //if (mTableProduceFinishedCount>0 && mTableProduceFinishedCount % 100 == 0)
                        //{
                        //    if (mc.Execute())
                        //    {
                        //        memo = string.Format("提交计算完工数第{1}到第{0}行，耗时：{2:#,##0.000}秒。", mTableProduceFinishedCount,mTableProduceFinishedCountBatch, (DateTime.Now - mStartTime1).TotalSeconds);
                        //        mStartTime1 = DateTime.Now;
                        //        MyRecord.Say(memo);
                        //    }
                        //    else
                        //    {
                        //        memo = string.Format("提交计算完工数第{1}到第{0}行，耗时：{2:#,##0.000}秒。错误！！！", mTableProduceFinishedCount, mTableProduceFinishedCountBatch, (DateTime.Now - mStartTime1).TotalSeconds);
                        //        mStartTime1 = DateTime.Now;
                        //        MyRecord.Say(memo);
                        //    }
                        //    mTableProduceFinishedCountBatch = mTableProduceFinishedCount;
                        //    mc.SQLCmdColl=null;
                        //    mc = new MyData.MyCommand();
                        //}
                        string RdsNo = Convert.ToString(r["RdsNo"]);
                        DateTime mStartTime = DateTime.Now;
                        //MyData.MyParameter mp = new MyData.MyParameter("@rdsno", RdsNo);
                        //string SQL2 = "Exec [_PMC_UpdateFinishAndPicking] @RdsNo";
                        //mc.Add(SQL2, string.Format("SQLUPDATE_{0}", RdsNo), mp);
                        //memo = string.Format("添加第{0}条，工程单号：{1}，加一条等一秒。", mTableProduceFinishedCount, RdsNo);
                        if (!ProduceFeedBackCalculator(RdsNo))
                        {
                            memo = string.Format("A计算完工数第{0}条，工程单号：{1}，不成功，耗时：{2:#,#0.00}秒。", mTableProduceFinishedCount, RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        else
                        {
                            memo = string.Format("A计算完工数第{0}条，工程单号：{1}，完成，耗时：{2:#,#0.00}秒。", mTableProduceFinishedCount, RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        MyRecord.Say(memo);
                        mTableProduceFinishedCount++;
                    }
                }
                else
                {
                    MyRecord.Say("没有获取到任何内容");
                }
                MyRecord.Say(string.Format("计算完毕，一共耗时：{0:#,##0.00}分钟，下次启动于：{1:HH:mm}", (DateTime.Now - NowTime).TotalMinutes, NowTime.AddHours(3)));
                MyRecord.Say("-----------------------计算完成。-------------------------------");
                Thread.Sleep(1000);
                ProduceFeedBackRuning = false;
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        bool ProduceFeedBackCalculator(string CurrentRdsNO)
        {
            try
            {
                MyData.MyCommand mc = new MyData.MyCommand();

                MyData.MyDataParameter mp = new MyData.MyDataParameter("@rdsno", CurrentRdsNO);
                string SQL = "Exec [dbo].[_PMC_UpdateFinishAndPicking] @RdsNo";
                return mc.Execute(SQL, mp);

            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
                return false;
            }
        }
        #endregion

        #region 批量计算完工数

        private int mCacluateProduceFeedBackTimes = 0;

        void ProduceFeedBackManaulLoder()
        {
            MyRecord.Say("批量计算完工数和领料数线程创建.......");
            Thread t = new Thread(ProduceFeedback);
            t.IsBackground = true;
            MyRecord.Say("计算完工数和领料数线程已经启动。");
            t.Start();
        }

        void ProduceFeedback()
        {
            try
            {
                string SQL = "";
                MyRecord.Say("---------------------启动计算完工数。------------------------------");
                DateTime NowTime = DateTime.Now;
                DateTime StartTime = NowTime.AddDays(-mCacluateProduceFeedBackTimes), StopTime = NowTime;
                mCacluateProduceFeedBackTimes = 0;
                ProduceFeedBackLastRunTime = NowTime;
                MyRecord.Say(string.Format("计算起始时间：{0:yy/MM/dd HH:mm}，结束时间：{1:yy/MM/dd HH:mm}", StartTime, StopTime));
                MyRecord.Say("定时计算——获取计算范围");
                SQL = @"
							Select RdsNo From moProduce Where InputDate Between @InputBegin And @InputEnd order by Inputdate Desc
						";
                MyData.MyDataTable mTableProduceFinished = new MyData.MyDataTable(SQL,
                    new MyData.MyDataParameter("@InputBegin", StartTime, MyData.MyDataParameter.MyDataType.DateTime),
                    new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime)
                    );
                if (_StopAll) return;
                if (mTableProduceFinished != null && mTableProduceFinished.MyRows.Count > 0)
                {
                    string memo = string.Format("读取了{0}条记录，下面开始计算....", mTableProduceFinished.MyRows.Count);
                    MyRecord.Say(memo);
                    int mTableProduceFinishedCount = 1;
                    var v = from a in mTableProduceFinished.MyRows
                            select a;
                    DateTime mStartTime1 = DateTime.Now;
                    MyData.MyCommand mc = new MyData.MyCommand();
                    foreach (MyData.MyDataRow r in v)
                    {
                        if (_StopAll) return;
                        string RdsNo = Convert.ToString(r["RdsNo"]);
                        DateTime mStartTime = DateTime.Now;
                        //MyData.MyParameter mp = new MyData.MyParameter("@rdsno", RdsNo);
                        //string SQL2 = "Exec [_PMC_UpdateFinishAndPicking] @RdsNo";
                        //mc.Add(SQL2, string.Format("SQLUPDATE_{0}", RdsNo), mp);
                        //memo = string.Format("添加第{0}条，工程单号：{1}，加一条等一秒。", mTableProduceFinishedCount, RdsNo);
                        if (!ProduceFeedBackCalculator(RdsNo))
                        {
                            memo = string.Format("B计算完工数第{0}条，工程单号：{1}，不成功，耗时：{2:#,#0.00}秒。", mTableProduceFinishedCount, RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        else
                        {
                            memo = string.Format("B计算完工数第{0}条，工程单号：{1}，完成，耗时：{2:#,#0.00}秒。", mTableProduceFinishedCount, RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        MyRecord.Say(memo);
                        mTableProduceFinishedCount++;
                    }
                }
                else
                {
                    MyRecord.Say("没有获取到任何内容");
                }
                MyRecord.Say(string.Format("计算完毕，一共耗时：{0:#,##0.00}分钟，下次启动于：{1:HH:mm}", (DateTime.Now - NowTime).TotalMinutes, NowTime.AddHours(3)));
                MyRecord.Say("-----------------------计算完成。-------------------------------");
                Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }
        #endregion

        #region 定时审核排程
        #region 每隔9分钟检查一次是否有30分钟外没有审核的已撤销排程

        void ConfirmPlanHalfHour()
        {
            Thread t = new Thread(ConfirmPlanRecordCheckByHalfHour);
            t.IsBackground = true;
            t.Start();
        }

        void ConfirmPlanRecordCheckByHalfHour()
        {
            MyRecord.Say("---------------------定时检查已撤销没提交的排程---------------------------------");
            DateTime bStartTime = DateTime.Now;
            DateTime NowTime = DateTime.Now, TimeSet = DateTime.MinValue;
            try
            {
                string SQL = "";
                MyRecord.Say("----刷新所有已撤销没提交的排程。");
                SQL = @"
Select a.[_ID] as ID,a.RdsNo,a.PlanBegin,b.[date]
  from [_PMC_ProdPlan] a Inner Join 
	   (Select Max([date]) as [date],zbid From [_PMC_ProdPlan_StatusRecorder] Where [Type] ='撤銷提交' Group by zbid) b On a.[_ID] = b.zbid
 Where a.Status < 2 And DateDiff(MI,b.[date],GetDate()) > 30";
                MyData.MyDataTable mLastPlan = new MyData.MyDataTable(SQL);
                if (_StopAll) return;
                if (mLastPlan != null && mLastPlan.MyRows.Count > 0)
                {
                    string memo = string.Format("----读取了{0}条记录，下面开始计算....", mLastPlan.MyRows.Count);
                    MyRecord.Say(memo);
                    int mLastPlanCount = 1;
                    foreach (MyData.MyDataRow r in mLastPlan.MyRows)
                    {
                        if (_StopAll) return;
                        int PlanID = r.IntValue("ID");
                        string PlanRdsNo = r.Value("RdsNo");
                        DateTime xStartTime = r.DateTimeValue("PlanBegin"), xCheckDate = r.DateTimeValue("date");
                        MyRecord.Say(string.Format("----审核第{0}条；排程号：{1}，撤销时间：{2:yy.MM.dd HH:mm}", mLastPlanCount, PlanRdsNo, xCheckDate));
                        if (xStartTime.Hour < 17)
                        {
                            TimeSet = xStartTime.Date.AddHours(20);
                        }
                        else
                        {
                            TimeSet = xStartTime.Date.AddDays(1).Date.AddHours(8);
                        }
                        DateTime mStartTime = DateTime.Now;
                        string xMemo = string.Format("自动审核撤销确认的排程，排程起始时间：{0:yy.MM.dd HH:mm}，设置结束时间：{1:yy.MM.dd HH:mm}", xStartTime, TimeSet);
                        MyRecord.Say(string.Format("----{0}", xMemo));
                        ConfirmAndCheckPlan(PlanID, TimeSet, xMemo);
                        MyRecord.Say(string.Format("----耗时：{0:#,#0.00000000}秒。", (DateTime.Now - mStartTime).TotalSeconds));
                        mLastPlanCount++;
                    }
                }
                else
                {
                    MyRecord.Say("----没有获取到任何内容。");
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
            MyRecord.Say(string.Format("----计算完毕，一共耗时：{0:#,##0.000000}秒。", (DateTime.Now - bStartTime).TotalSeconds));
            MyRecord.Say("-----------------------定时检查已撤销没提交的排程 --完毕-------------------------------");
        }

        #endregion
        void ConfirmProcessPlanLoader()
        {
            MyRecord.Say("开启定时审核排程线程..........");
            Thread t = new Thread(ConfirmProcessPlanRecord);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("定时审核排程线程已经开启。");
        }

        void ConfirmProcessPlanRecord()
        {
            try
            {
                string DeptSQL = "", ProcSQL = "";
                MyRecord.Say("---------------------启动定时审核排程。------------------------------");
                DateTime bStartTime = DateTime.Now;
                DateTime NowTime = DateTime.Now, TimeBegin = DateTime.MinValue, TimeEnd = DateTime.MinValue, TimeSet = DateTime.MinValue;
                string xMemo = string.Empty, OnWorkFieldName = string.Empty, OffWorkFieldName = string.Empty;
                if (NowTime.Hour < 12 && NowTime.Hour > 5)
                {
                    OnWorkFieldName = "OnWorkTime1"; OffWorkFieldName = "OffWorkTime1";
                    xMemo = "白班自动审核生产计划。";
                }
                else
                {
                    OnWorkFieldName = "OnWorkTime2"; OffWorkFieldName = "OffWorkTime2";
                    xMemo = "夜班自动审核生产计划。";
                }
                MyRecord.Say(string.Format("逐部门查询各个工序最后一份排程。OnWorkFieldName={0}，OffWorkFieldName={1}，{2}", OnWorkFieldName, OffWorkFieldName, xMemo));
                DeptSQL = @"Select * from pbDept Where Not (OnWorkTime1 is Null And OnWorkTime2 is Null) Order by FullSortID";
                MyData.MyDataTable mdDepartment = new MyData.MyDataTable(DeptSQL);
                if (mdDepartment.MyRows.IsNotEmptySet())
                {
                    MyRecord.Say(string.Format("各个部门已经获取，共{0}部门。", mdDepartment.MyRows.Count));
                    int dpCount = 1;
                    foreach (var deptItem in mdDepartment.MyRows)
                    {
                        ProcSQL = @"
Select Max(a.[_id]) as ID,RdsNo =(Select RdsNo From _PMC_ProdPlan r Where r.[_id]=Max(a.[_id])),a.Process,ProcessName=(Select Name from moProcedure Where Code = a.Process)
  from [_PMC_ProdPlan] a
 Where a.PlanBegin Between @TimeBgein And @TimeEnd And PlanEnd is Null And
	   Not Exists(Select Top 1 _ID From [_PMC_ProdPlan] b Where b.PlanBegin Between @TimeBgein And @TimeEnd And
																b.PlanEnd is Not Null And a.Process=b.Process And a.Department=b.Department)
 Group by a.Process,a.Department
 Having a.Department = @DeptID
";
                        TimeBegin = NowTime.Date.Add((deptItem.DateTimeValue(OnWorkFieldName) - DateTime.Parse("2000-01-01").Date));
                        TimeEnd = NowTime.Date.Add((deptItem.DateTimeValue(OffWorkFieldName).AddHours(-2) - DateTime.Parse("2000-01-01").Date));
                        TimeSet = NowTime.Date.Add((deptItem.DateTimeValue(OffWorkFieldName) - DateTime.Parse("2000-01-01").Date));
                        MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
						{
							new MyData.MyDataParameter("@TimeBgein",TimeBegin, MyData.MyDataParameter.MyDataType.DateTime),
							new MyData.MyDataParameter("@TimeEnd",TimeEnd, MyData.MyDataParameter.MyDataType.DateTime),
							new MyData.MyDataParameter("@DeptID",deptItem.IntValue("_ID"), MyData.MyDataParameter.MyDataType.Int)
						};
                        MyRecord.Say(string.Format("{0}.{1}，排程起始：{2}到：{3}止，结束时间：{4}", dpCount, deptItem.Value("Name"), TimeBegin, TimeEnd, TimeSet));
                        using (MyData.MyDataTable mLastPlan = new MyData.MyDataTable(ProcSQL, mps))
                        {
                            if (mLastPlan.MyRows.IsNotEmptySet())
                            {
                                MyRecord.Say(string.Format("1.1.{2}，读取了{0}条记录，下班时间：{1:yyyy-MM-dd HH:mm}，下面开始计算....", mLastPlan.MyRows.Count, TimeSet, dpCount));
                                int mLastPlanCount = 1;
                                foreach (MyData.MyDataRow r in mLastPlan.MyRows)
                                {
                                    int PlanID = r.IntValue("ID");
                                    MyRecord.Say(string.Format("1.1.{0}.{1}，工序：{2}，PlanID={3}，排程：{4}", dpCount, mLastPlanCount, r.Value("ProcessName"), PlanID, r.Value("RdsNo")));
                                    DateTime mStartTime = DateTime.Now;
                                    ConfirmAndCheckPlan(PlanID, TimeSet, xMemo);
                                    MyRecord.Say(string.Format("耗时：{0:#,#0.00}秒。", (DateTime.Now - mStartTime).TotalSeconds));
                                    mLastPlanCount++;
                                }
                            }
                            else
                            {
                                MyRecord.Say("没有获取到任何内容");
                            }
                        }
                        dpCount++;
                    }
                    MyRecord.Say(string.Format("计算完毕，一共耗时：{0:#,##0.00}分钟。", (DateTime.Now - bStartTime).TotalSeconds));
                    MyRecord.Say("-----------------------审核完成。-------------------------------");
                }
                else
                {
                    MyRecord.Say("没有找到要审核的部门。");
                    MyRecord.Say("-----------------------审核出错。-------------------------------");
                }

            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        bool ConfirmAndCheckPlan(int xPlanID, DateTime PlanEndSet, string Memo)
        {
            string SQL = @"
SET NOCOUNT ON
Declare @PlanEnd DateTime
Declare @Status int
Declare @rdsno Varchar(40)
Select @PlanEnd=PlanEnd,@Status=Status,@rdsno=RdsNo From [_PMC_ProdPlan] Where [_ID]=@zbid
if (@PlanEnd is Null)
Begin
	Update [_PMC_ProdPlan] Set PlanEnd=@PlanEndSet Where _ID=@zbid
End
if (@Status < 1)
Begin
	Insert Into [_PMC_ProdPlan_StatusRecorder]
			(zbid,Author,[state],memo,rdsno,type,typeid,CheckIn,CheckOut)
	Values(@zbid,'系統審核',+1,@Memo,@rdsno,'審核',2,1,1)
End
Update [_PMC_ProdPlan] Set Status=2,Checker='系統審核',CheckDate=GetDate() Where [_ID]=@zbid
Insert Into [_PMC_ProdPlan_StatusRecorder]
		(zbid,Author,[state],memo,rdsno,type,typeid,CheckIn,CheckOut)
Values(@zbid,'系統審核',+1,@Memo,@rdsno,'確認',2,1,1)
SET NOCOUNT OFF
";
            MyData.MyCommand mc = new MyData.MyCommand();
            return mc.Execute(SQL,
                new MyData.MyDataParameter("@zbid", xPlanID, MyData.MyDataParameter.MyDataType.Int),
                new MyData.MyDataParameter("@PlanEndSet", PlanEndSet, MyData.MyDataParameter.MyDataType.DateTime),
                new MyData.MyDataParameter("@Memo", Memo)
                );
        }

        #endregion

        #region 发送昨天未审核工单

        void SendProduceDayReportLoder()
        {
            MyRecord.Say("开启昨日未审核工单..........");
            Thread t = new Thread(ProduceDayReport);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启昨日未审核工单..完成。");
        }

        void ProduceDayReport()
        {
            try
            {
                MyRecord.Say("-----------------昨日未审核工单-------------------------");
                MyRecord.Say("生成内容");

                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{4}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容，（输入日期在{0:yy/MM/dd HH:mm} - {1:yy/MM/dd HH:mm}之间）未审核工单列表。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请确认是否需要审核和未审核原因。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	客户
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	产品编号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	产品料号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	版号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	订单编号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	订单数量
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单数量
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	输入人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	输入时间
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单性质
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	工单交期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	交期备注
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	刻交期人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	是否入库
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	补单原单号
	</TD>
	</TR>
	{2}
</TBODY></TABLE></FONT>
</DIV>
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{3:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
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
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	{15}
	</TD>
	</TR>
";

                string SQL = @"
Select *,ProductName =(Select Name from pbProduct p Where p.Code=a.Code),
		 PropertyName=(Select Name from moProdProperty pp Where pp.Code=a.Property)
 from moProduce a
Where isNull(Status,0) = 0 And StockDate is Null And FinishDate is Null
  And a.InputDate Between @DateBegin And @DateEnd
";
                DateTime NowTime = DateTime.Now;
                DateTime BDate = NowTime.AddYears(-1).Date, EDate = NowTime.Date.AddHours(NowTime.Hour);
                string brs = "";
                MyRecord.Say("后台计算");
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL,
                    new MyData.MyDataParameter("@DateBegin", BDate, MyData.MyDataParameter.MyDataType.DateTime),
                    new MyData.MyDataParameter("@DateEnd", EDate, MyData.MyDataParameter.MyDataType.DateTime)
                    ))
                {
                    MyRecord.Say("工单");
                    if (md != null && md.MyRows.Count > 0)
                    {
                        foreach (var ri in md.MyRows)
                        {
                            string spNumb = Convert.ToString(ri["pNumb"]), soNumb = Convert.ToString(ri["oNumb"]);
                            int pNumb = 0; if (!int.TryParse(spNumb, out pNumb)) pNumb = 0;
                            int oNumb = 0; if (!int.TryParse(soNumb, out oNumb)) oNumb = 0;
                            brs += string.Format(br, Convert.ToString(ri["RdsNo"]),
                                                    Convert.ToString(ri["Custid"]),
                                                    Convert.ToString(ri["Code"]),
                                                    Convert.ToString(ri["ProductName"]),
                                                    Convert.ToString(ri["PlateCode"]),
                                                    Convert.ToString(ri["OrderNo"]),
                                                    oNumb,
                                                    pNumb,
                                                    Convert.ToString(ri["Inputer"]),
                                                    string.Format("{0:yy/MM/dd HH:mm}", Convert.ToDateTime(ri["InputDate"])),
                                                    Convert.ToString(ri["PropertyName"]),
                                                    string.Format("{0:yy/MM/dd HH:mm}", Convert.ToDateTime(ri["pDeliver"])),
                                                    Convert.ToString(ri["SRemark"]),
                                                    Convert.ToString(ri["RMan"]),
                                                    Convert.ToInt32(ri["StopStock"]) == 0 ? "" : "无需入库",
                                                    Convert.ToString(ri["RenewRdsNo"])
                                                    );
                        }
                        MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", md.Rows.Count));
                        MyRecord.Say("创建SendMail。");
                        MyBase.SendMail sm = new MyBase.SendMail();
                        MyRecord.Say("加载邮件内容。");
                        sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, BDate, EDate, brs, NowTime, MyBase.CompanyTitle));
                        sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_昨日未审核工单提醒。", NowTime, MyBase.CompanyTitle));
                        //string mailto = ConfigurationManager.AppSettings["ProduceMailTo"], mailcc = ConfigurationManager.AppSettings["ProduceMailCC"];
                        //MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        //sm.MailTo = mailto; //"my18@my.imedia.com.tw,xang@my.imedia.com.tw,lghua@my.imedia.com.tw,my64@my.imedia.com.tw";
                        //sm.MailCC = mailcc; //"jane123@my.imedia.com.tw,lwy@my.imedia.com.tw,my80@my.imedia.com.tw";
                        ////sm.MailTo = "my80@my.imedia.com.tw";

                        MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("Produce");
                        MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                        sm.MailTo = mAddress.MailTo;
                        sm.MailCC = mAddress.MailCC;
                        //sm.MailTo = "my80@my.imedia.com.tw";

                        MyRecord.Say("发送邮件。");
                        sm.SendOut();
                        MyRecord.Say("已经发送。");
                    }
                    else
                    {
                        MyRecord.Say("没有找到资料。");
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

        #region 定时清理生产排程记录
        bool PlanRecordCleanRuning = false;

        void PlanRecordCleanLoder()
        {
            MyRecord.Say("定时清理生产排程记录..........");
            Thread t = new Thread(PlanRecordClean);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("定时清理生产排程记录已经开启。");
        }

        void PlanRecordClean()
        {
            try
            {
                DateTime rxtime = DateTime.Now;
                while (ProduceFeedBackRuning || CheckStockRecordRunning)
                {
                    if ((DateTime.Now - rxtime).TotalMinutes > 20) return;
                    Thread.Sleep(5000);
                };

                PlanRecordCleanRuning = true;
                string SQL = "";
                MyRecord.Say("---------------------定时清理生产排程记录------------------------------");
                DateTime NowTime = DateTime.Now;
                MyRecord.Say("1.定时清理3天之前，所有未审核/未确认的生产排程。");

                SQL = @"
Declare @NowDate DateTime,@LastDate DateTime
Declare @WeekDay int
Set @NowDate = GetDate()
Set @WeekDay=DatePart(DW,@NowDate)
Select @LastDate = (Case When @WeekDay <= 4 Then DateAdd([dd],-4,@NowDate) Else DateAdd([dd],-3,@NowDate) End)
Select [_ID],RdsNo from _PMC_ProdPlan Where IsNull([Status],0) < 2 And PlanBegin < @LastDate";
                MyData.MyDataTable mTableUnCheckPlan = new MyData.MyDataTable(SQL);
                if (_StopAll) return;
                if (mTableUnCheckPlan != null && mTableUnCheckPlan.MyRows.Count > 0)
                {
                    string memo = string.Format("读取了{0}条记录，下面开始刪除....", mTableUnCheckPlan.MyRows.Count);
                    MyRecord.Say(memo);
                    int mTableProduceFinishedCount = 1;
                    foreach (MyData.MyDataRow r in mTableUnCheckPlan.MyRows)
                    {
                        if (_StopAll) return;
                        string RdsNo = Convert.ToString(r["RdsNo"]);
                        DateTime mStartTime = DateTime.Now;
                        MyData.MyDataParameter mpid = new MyData.MyDataParameter("@ID", r["_ID"], MyData.MyDataParameter.MyDataType.Int);
                        string SQLDetial = @"Delete from [_PMC_ProdPlan_List] Where zbid=@ID", SQLMain = @"Delete from [_PMC_ProdPlan] Where [_ID]=@ID";
                        MyData.MyCommand DeleteCMD = new MyData.MyCommand();
                        DeleteCMD.Add(SQLDetial, "DeleteDetail", mpid);
                        DeleteCMD.Add(SQLMain, "DeleteMain", mpid);
                        if (DeleteCMD.Execute())
                        {
                            memo = string.Format("删除第{0}条，排程单号：{1}，成功，耗时：{2:#,#0.00}秒。", mTableProduceFinishedCount, RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        else
                        {
                            memo = string.Format("删除第{0}条，排程单号：{1}，完成，耗时：{2:#,#0.00}秒。", mTableProduceFinishedCount, RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        MyRecord.Say(memo);
                        mTableProduceFinishedCount++;
                    }
                }
                else
                {
                    MyRecord.Say("没有获取到任何内容");
                }
                MyRecord.Say(string.Format("清理完毕，下面进行第二项，一共耗时：{0:#,##0.00}分钟，下次启动于：{1:HH:mm}", (DateTime.Now - NowTime).TotalMinutes, NowTime.AddHours(3)));

                MyRecord.Say("2.定时清理3天之前，所有已审核的生产排程。进行关闭。");
                SQL = @"Declare @NowDate DateTime,@LastDate DateTime
Declare @WeekDay int
Set @NowDate = GetDate()
Set @WeekDay=DatePart(DW,@NowDate)
Select @LastDate = (Case When @WeekDay <= 4 Then DateAdd([dd],-4,@NowDate) Else DateAdd([dd],-3,@NowDate) End)
Select [_ID],RdsNo from _PMC_ProdPlan Where IsNull([Status],0) = 2 And PlanBegin < @LastDate";
                MyData.MyDataTable mTableCheckedPlan = new MyData.MyDataTable(SQL);
                if (_StopAll) return;
                if (mTableCheckedPlan != null && mTableCheckedPlan.MyRows.Count > 0)
                {
                    string memo = string.Format("读取了{0}条记录，下面开始计算....", mTableCheckedPlan.MyRows.Count);
                    MyRecord.Say(memo);
                    int mTableProduceFinishedCount = 1;
                    foreach (MyData.MyDataRow r in mTableCheckedPlan.MyRows)
                    {
                        if (_StopAll) return;
                        DateTime mStartTime = DateTime.Now;
                        string RdsNo = Convert.ToString(r["RdsNo"]);
                        MyData.MyDataParameter mpRdsNo = new MyData.MyDataParameter("@rdsno", RdsNo);
                        MyData.MyDataParameter mpid = new MyData.MyDataParameter("@ID", r["_ID"], MyData.MyDataParameter.MyDataType.Int);
                        MyData.MyDataParameter mpEndDate = new MyData.MyDataParameter("@PlanEnd", r["PlanEnd"], MyData.MyDataParameter.MyDataType.DateTime);
                        string SQLDetial = @"Delete from [_PMC_ProdPlan_List] Where zbid = @ID And Bdd > DateAdd(HH,3,@PlanEnd)";
                        string SQLMain = @"Update [_PMC_ProdPlan] Set Status=3 Where [_ID]=@ID";
                        string SQLStatus = @"Insert Into [_PMC_ProdPlan_StatusRecorder](zbid,Author,[state],memo,rdsno,type,typeid,CheckIn,CheckOut) Values(@ID,'自動審核',1,'關閉排程，到 ' + Convert(VarChar(100),DateAdd(HH,3,@PlanEnd),120),@rdsno,'關閉',2,1,1)";
                        MyData.MyDataParameter[] mpss = new MyData.MyDataParameter[]
						{
							mpRdsNo,
							mpid,
							mpEndDate
						};
                        MyData.MyCommand doCMD = new MyData.MyCommand();
                        doCMD.Add(SQLDetial, "Delete", mpss);
                        doCMD.Add(SQLMain, "UpdateMain", mpss);
                        doCMD.Add(SQLStatus, "Status", mpss);
                        if (doCMD.Execute())
                        {
                            memo = string.Format("設置第{0}条，排程单号：{1}，成功，耗时：{2:#,#0.00}秒。", mTableProduceFinishedCount, RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        else
                        {
                            memo = string.Format("設置第{0}条，排程单号：{1}，完成，耗时：{2:#,#0.00}秒。", mTableProduceFinishedCount, RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        MyRecord.Say(memo);
                        mTableProduceFinishedCount++;
                    }
                }
                else
                {
                    MyRecord.Say("没有获取到任何内容");
                }
                MyRecord.Say(string.Format("清理完毕，一共耗时：{0:#,##0.00}分钟，下次启动于：{1:HH:mm}", (DateTime.Now - NowTime).TotalMinutes, NowTime.AddHours(3)));
                MyRecord.Say("-----------------------清理完毕。-------------------------------");
                PlanRecordCleanRuning = false;
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }
        #endregion

        #region 发送每日8D未回复8D报告

        void ProblemSolving8DReportLoder()
        {
            MyRecord.Say("开启昨日回复8D报告..........");
            Thread t = new Thread(ProblemSolving8D);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启昨日未回复8D报告..完成。");
        }

        void ProblemSolving8D()
        {
            try
            {
                MyRecord.Say("-----------------昨日未回复8D报告-------------------------");
                MyRecord.Say("生成内容");
                #region 邮件体

                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{4}ERP系统提示您：</FONT></DIV>
{0}
{1}
{2}
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{3:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");

                string body1 = @"
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容，为8H前（输入日期在{0:yy/MM/dd HH:mm}之前）未回复/未审核/被退件的8D报告列表。(每日7点更新统计)</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请相关单位主管及时处理，CQE将持续追踪进度。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	类型
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	发起人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	异常项目
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	责任部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	指定答辩人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	发起日期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	答复日期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	审核日期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	逾期时间
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	当前状态
	</TD>
	</TR>
	{1}
</TBODY></TABLE></FONT>
</DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </FONT></DIV>
";
                string body2 = @"
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请品保注意以下内容，为8H前（输入日期在{0:yy/MM/dd HH:mm}之前）已审核待品保处理/确认的8D报告列表。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	类型
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	发起人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	异常项目
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	责任部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	指定答辩人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	发起日期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	回覆日期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	审核日期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	逾期时间
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	当前状态
	</TD>
	</TR>
	{1}
</TBODY></TABLE></FONT>
</DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </FONT></DIV>
				";
                string body3 = @"
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请品保注意以下内容，为即将到预期结案日期的8D报告列表，请尽快追踪结案。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	类型
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	发起人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	异常项目
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	责任部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	指定答辩人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	审核日期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	确认日期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	预计结案日期
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	到期时间
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	当前状态
	</TD>
	</TR>
	{0}
</TBODY></TABLE></FONT>
</DIV>
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
";
                #endregion
                #region 邮件变体
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
	</TR>
";
                #endregion

                string SQL = @"
Select *,DepartmentName=(Select Top 1 name from pbDept Where pbDept.Code=a.DepartmentCode),
		 UserName=(Select Top 1 Name from [_HR_Employee] b Where b.Code=a.SendToUserCode),
		 TypeName=(Case When isNull(a.Type,0)=1 Then '廠內'
					  When isNull(a.Type,0)=2 Then '廠外客訴/客退'
					  When isNull(a.Type,0)=3 Then '供應商' Else '' END),
		 FinishTypeName=(Case When isNull(a.FinishType,0)=1 Then '結案'
					  When isNull(a.FinishType,0)=2 Then '繼續跟催'
					  When isNull(a.FinishType,0)=3 Then '重擬對策再做跟催' Else '處理中' END),
		 StatusName=(Select (Case When a.Status = 0 And Not a.ModifyDate is Null Then '待審核' Else Name End) from [_SY_Status] Where [Type]='8D' And StatusID=isNull(a.Status,0))
from [_QE_ProblemSolving8D] a Where FinalizDate is Null And InputDate <= @DateEnd Order By a.Status,a.FinishType Desc,a.InputDate
";
                DateTime NowTime = DateTime.Now;
                DateTime EDate = NowTime.Date.AddHours(8);
                string brs = "", brt = "", brl = "";
                MyRecord.Say("后台计算");
                //类型，单号,发起人,问题描述,责任部门,指定答辩人,发起日期,逾期时间,当前状态
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@DateEnd", EDate, MyData.MyDataParameter.MyDataType.DateTime)))
                {
                    MyRecord.Say("读出8D报告，创建表格内容");
                    if (md != null && md.MyRows.Count > 0)
                    {
                        MyRecord.Say(string.Format("共：{0}", md.MyRows.Count));
                        int iCount = 0;
                        foreach (var ri in md.MyRows)
                        {
                            try
                            {
                                int statusid = Convert.ToInt32(ri["Status"]), finishtypeid = Convert.ToInt32(ri["FinishType"]);
                                string statusname = "待回复", overtimeword;
                                DateTime modifydate = Convert.ToDateTime(ri["ModifyDate"]), checkdate = Convert.ToDateTime(ri["CheckDate"]);
                                iCount += 1;
                                MyRecord.Say(string.Format("第：{0}条，单号：{1}", iCount, ri["RdsNo"]));
                                if (statusid <= 0 || (statusid == 1 && finishtypeid == 3))
                                {
                                    if (finishtypeid == 3)
                                        statusname = "被退件！需重新回覆。";
                                    else
                                    {
                                        if (modifydate > DateTime.MinValue)
                                            statusname = "等待主管审核。";
                                        else
                                            statusname = "未回覆。";
                                    }
                                    //类型，单号,发起人,问题描述,责任部门,指定答辩人,发起日期,回覆日期，审核日期，逾期时间,当前状态
                                    DateTime startdate = Convert.ToDateTime(ri["InputDate"]);
                                    overtimeword = (NowTime - startdate).TotalHours > 24 ? string.Format("{0:#.#}天", (NowTime - startdate).TotalDays) : string.Format("{0:0}小时", (NowTime - startdate).TotalHours);
                                    brs += string.Format(br, Convert.ToString(ri["TypeName"]).Substring(0, 2),
                                                             Convert.ToString(ri["RdsNo"]),
                                                             Convert.ToString(ri["Inputer"]),
                                                             Convert.ToString(ri["ProjName"]),
                                                             Convert.ToString(ri["DepartmentName"]),
                                                             Convert.ToString(ri["UserName"]),
                                                             startdate.ToString("yy/MM/dd HH:mm"),
                                                             modifydate > DateTime.MinValue ? modifydate.ToString("yy/MM/dd HH:mm") : "",
                                                             checkdate > DateTime.MinValue ? checkdate.ToString("yy/MM/dd HH:mm") : "",
                                                             overtimeword,
                                                             statusname
                                                            );

                                }
                                else
                                {
                                    DateTime finishdate = Convert.ToDateTime(ri["FinishDate"]), overdate = Convert.ToDateTime(ri["D8_FinishDate"]);
                                    statusname = Convert.ToString(ri["StatusName"]);
                                    if (finishdate <= DateTime.MinValue)
                                    {
                                        //类型，单号,发起人,问题描述,责任部门,指定答辩人,发起日期,回覆日期，审核日期，逾期时间,当前状态
                                        DateTime startdate = checkdate;
                                        overtimeword = (NowTime - startdate).TotalHours > 24 ? string.Format("{0:#.#}天", (NowTime - startdate).TotalDays) : string.Format("{0:0}小时", (NowTime - startdate).TotalHours);
                                        brl += string.Format(br, Convert.ToString(ri["TypeName"]).Substring(0, 2),
                                                                 Convert.ToString(ri["RdsNo"]),
                                                                 Convert.ToString(ri["Inputer"]),
                                                                 Convert.ToString(ri["ProjName"]),
                                                                 Convert.ToString(ri["DepartmentName"]),
                                                                 Convert.ToString(ri["UserName"]),
                                                                 Convert.ToDateTime(ri["D1_Date"]).ToString("yy/MM/dd HH:mm"),
                                                                 modifydate > DateTime.MinValue ? modifydate.ToString("yy/MM/dd HH:mm") : "",
                                                                 checkdate > DateTime.MinValue ? checkdate.ToString("yy/MM/dd HH:mm") : "",
                                                                 overtimeword,
                                                                 statusname
                                                                );
                                    }
                                    else if (finishdate > DateTime.MinValue && (overdate - NowTime).TotalDays < 3)
                                    {
                                        //类型，单号,发起人,问题描述,责任部门,指定答辩人,审核日期,处理日期，预计结案日期，逾期时间,当前状态
                                        overtimeword = (overdate - NowTime).TotalHours > 24 ? string.Format("{0:#.#}天", (overdate - NowTime).TotalDays) : string.Format("{0:0}小时", (overdate - NowTime).TotalHours);
                                        brt += string.Format(br, Convert.ToString(ri["TypeName"]).Substring(0, 2),
                                                                 Convert.ToString(ri["RdsNo"]),
                                                                 Convert.ToString(ri["Inputer"]),
                                                                 Convert.ToString(ri["ProjName"]),
                                                                 Convert.ToString(ri["DepartmentName"]),
                                                                 Convert.ToString(ri["UserName"]),
                                                                 overdate.ToString("yy/MM/dd HH:mm"),
                                                                 Convert.ToDateTime(ri["D7_Date"]) > DateTime.MinValue ? Convert.ToDateTime(ri["D7_Date"]).ToString("yy/MM/dd HH:mm") : "",
                                                                 Convert.ToDateTime(ri["D8_FinishDate"]) > DateTime.MinValue ? Convert.ToDateTime(ri["D8_FinishDate"]).ToString("yy/MM/dd HH:mm") : "",
                                                                 overtimeword,
                                                                 statusname
                                                                );
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MyRecord.Say(ex);
                            }
                        }
                        MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", md.Rows.Count));
                        MyRecord.Say("创建SendMail。");
                        MyBase.SendMail sm = new MyBase.SendMail();
                        MyRecord.Say("加载邮件内容。");
                        string sb1 = brs != null && brs.Length > 0 ? string.Format(body1, EDate, brs) : string.Empty;
                        string sb2 = brs != null && brl.Length > 0 ? string.Format(body2, EDate, brl) : string.Empty;
                        string sb3 = brs != null && brt.Length > 0 ? string.Format(body3, brt) : string.Empty;
                        sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, sb1, sb2, sb3, NowTime, MyBase.CompanyTitle));
                        sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_8D报告回覆进度追踪统计。", NowTime, MyBase.CompanyTitle));
                        //string mailto = ConfigurationManager.AppSettings["8DMailTo"], mailcc = ConfigurationManager.AppSettings["8DMailCC"];
                        //MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        //sm.MailTo = mailto;
                        //sm.MailCC = mailcc;
                        ////sm.MailTo = "my80@my.imedia.com.tw";

                        MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("8D");
                        MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                        sm.MailTo = mAddress.MailTo;
                        sm.MailCC = mAddress.MailCC;
                        //sm.MailTo = "my80@my.imedia.com.tw";

                        MyRecord.Say("发送邮件。");
                        sm.SendOut();
                        MyRecord.Say("已经发送。");
                    }
                    else
                    {
                        MyRecord.Say("没有找到资料。");
                        MyRecord.Say("创建SendMail。");
                        MyBase.SendMail sm = new MyBase.SendMail();
                        MyRecord.Say("加载邮件内容。");
                        string sb3 = MyConvert.ZH_TW(@"<DIV><FONT color=#FFF90051 size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;没有未回复的8D报告，非常棒！</STRONG></FONT></DIV>");
                        sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, "", "", sb3, NowTime, MyBase.CompanyTitle));
                        sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_8D报告回覆进度追踪统计。", NowTime, MyBase.CompanyTitle));

                        //string mailto = ConfigurationManager.AppSettings["8DMailTo"], mailcc = ConfigurationManager.AppSettings["8DMailCC"];
                        //MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        //sm.MailTo = mailto;
                        //sm.MailCC = mailcc;
                        ////sm.MailTo = "my80@my.imedia.com.tw";

                        MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("8D");
                        MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                        sm.MailTo = mAddress.MailTo;
                        sm.MailCC = mAddress.MailCC;
                        //sm.MailTo = "my80@my.imedia.com.tw";

                        MyRecord.Say("发送邮件。");
                        sm.SendOut();
                        MyRecord.Say("已经发送。");
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

        #region 整批计算完工数(暂时没启用)
        void UpdateProduceNoteFnishedNumberLoader()
        {
            MyRecord.Say("开始逐笔计算完工数和领料数。");
            Thread t = new Thread(UpdateProduceNoteFnishedNumber);
            t.IsBackground = true;
            MyRecord.Say("定时计算完工数线程已经开启。");
            t.Start();
        }

        void UpdateProduceNoteFnishedNumber()
        {
            try
            {
                MyRecord.Say("---------------------计算开始------------------------------");
                DateTime NowTime = DateTime.Now;
                DateTime StartTime = Convert.ToDateTime("2014-06-01").Date, StopTime = NowTime.Date.AddHours(NowTime.Hour);
                MyRecord.Say(string.Format("计算起始时间：{0:yy/MM/dd HH:mm}，结束时间：{1:yy/MM/dd HH:mm}", StartTime, StopTime));
                MyRecord.Say("定时计算——获取计算范围");
                string SQL = @"Select Distinct RdsNo from moProduce Where InputDate Between @InputBegin And @InputEnd Order by RdsNo";
                MyData.MyDataTable mTableProduceFinished = new MyData.MyDataTable(SQL,
                    new MyData.MyDataParameter("@InputBegin", StartTime, MyData.MyDataParameter.MyDataType.DateTime),
                    new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime)
                    );
                if (_StopAll) return;
                if (mTableProduceFinished != null && mTableProduceFinished.MyRows.Count > 0)
                {
                    string memo = string.Format("读取了{0}条记录，下面开始计算....", mTableProduceFinished.MyRows.Count);
                    MyRecord.Say(memo);
                    foreach (MyData.MyDataRow r in mTableProduceFinished.MyRows)
                    {
                        if (_StopAll) return;
                        string RdsNo = Convert.ToString(r["RdsNo"]);
                        DateTime mStartTime = DateTime.Now;
                        if (UpdateProduce(RdsNo))
                        {
                            memo = string.Format("X工程单号：{0}，完成，耗时：{1:#,#0.00}秒。", RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        else
                        {
                            memo = string.Format("X工程单号：{0}，不成功，耗时：{1:#,#0.00}秒。", RdsNo, (DateTime.Now - mStartTime).TotalSeconds);
                        }
                        MyRecord.Say(memo);
                    }
                }
                else
                {
                    MyRecord.Say("没有获取到任何内容");
                }
                MyRecord.Say(string.Format("计算完毕，一共耗时：{0:#,##0.00}分钟，下次启动于：{1:HH:mm}", (DateTime.Now - NowTime).TotalMinutes, NowTime.AddHours(3)));
                MyRecord.Say("-----------------------计算完成。-------------------------------");
                Thread.Sleep(1000);
                ProduceFeedBackRuning = false;
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        bool UpdateProduce(string CurrentRdsNO)
        {
            try
            {
                MyData.MyCommand mc = new MyData.MyCommand();
                MyData.MyDataParameter mp = new MyData.MyDataParameter("@rdsno", CurrentRdsNO);
                string SQL = "Exec [dbo].[_PMC_UpdateFinishAndPicking] @RdsNo";
                return mc.Execute(SQL, mp);
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
                return false;
            }
        }
        #endregion

        #region 发送每日未结束维修申请

        void MachineRepairReportLoder()
        {
            MyRecord.Say("开启未结束维修申请报告..........");
            Thread t = new Thread(MachineRepairReport);
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启未结束维修申请报告..完成。");
        }

        void MachineRepairReport()
        {
            try
            {
                MachineRepairReportChecker();
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }

            try
            {
                MyRecord.Say("-----------------未结束维修申请-------------------------");
                MyRecord.Say("生成内容");

                string body = LCStr(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{3}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容，为未结案的机台维修申请单列表。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请相关单位主管及时处理。黄色条目为：已经超出要求维修时间，还未维修的申请单，请特别注意！</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	单号
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	部门
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	故障程度
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	紧急程度
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	申请人
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	申请时间
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	机台名称
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	故障部位
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	故障描述
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	要求完成时间
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	处理状态
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	检修时间
	</TD>
	</TR>
	{1}
</TBODY></TABLE></FONT>
</DIV>
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{2:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");
                string br = @"
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{1}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{2}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{3}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{4}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{5}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{6}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{7}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{8}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{9}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{10}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{11}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {0}"" align=center >
	{12}
	</TD>
	</TR>
";

                string SQL = @"
Select a.*,b.Name as DepartmentName,c.Name as ProcessName,d.Name as LevelName,e.Name as ResonTitleName,f.name as RxTitleName,g.Name as ResultTypeName,h.name as FinishTypeName,
	   StatusName=(Select Name from [_SY_Status] Where [Type]='EM' And StatusID=isNull(a.Status,0)),
	   UrgentName=(Case When Urgent=0 Then '普通' When Urgent=1 Then '紧急' When Urgent=2 Then '特急' Else '' End)
from [_GS_EquipmentRepair] a Left Outer Join pbDept b On a.DepartmentCode=b.Code
							 Left Outer Join moProcedure c On a.ProcessCode=c.Code
							 Left Outer Join [_BS_Fault_Level] d ON d.Code=a.Level
							 Left Outer Join [_BS_Fault_Reason] e ON e.Code=a.ReasonTitle
							 Left Outer Join [_BS_Fault_RxTitle] f ON f.Code=a.RxTitle
							 Left Outer Join [_BS_Fault_ResultType] g ON g.Code=a.ResultType
							 Left Outer Join [_BS_Fault_FinishType] h ON h.Code=a.FinishType
Where a.Type='Machine' And a.CreateDate < @DateEnd And a.Status Between 0 And 3
Order by a.[_id]
";
                DateTime NowTime = DateTime.Now;
                DateTime EDate = NowTime.Date.AddDays(1).AddMilliseconds(-10);
                string brs = "";
                MyRecord.Say("后台计算");
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@DateEnd", EDate, MyData.MyDataParameter.MyDataType.DateTime)))
                {
                    MyRecord.Say("读出维修申请单，创建表格内容");
                    if (md != null && md.MyRows.Count > 0)
                    {
                        foreach (var ri in md.MyRows)
                        {
                            int statusid = Convert.ToInt32(ri["Status"]);
                            DateTime RequestFinishDate = Convert.ToDateTime(ri["RequestFinishDate"]), InspectionDate = Convert.ToDateTime(ri["InspectionDate"]);
                            string bkColor = "transparent";
                            if (statusid < 2 && (NowTime - RequestFinishDate).TotalHours > 0) bkColor = "yellow";
                            //部门,故障程度,紧急程度,申请人,申请时间,机台名称,故障部位,故障描述,要求完成时间,处理状态,检修时间
                            brs += string.Format(br, bkColor,
                                                     Convert.ToString(ri["RdsNo"]),
                                                     Convert.ToString(ri["DepartmentName"]),
                                                     Convert.ToString(ri["LevelName"]),
                                                     Convert.ToString(ri["UrgentName"]),
                                                     Convert.ToString(ri["Inputer"]),
                                                     Convert.ToDateTime(ri["InputDate"]).ToString("yy/MM/dd HH:mm"),
                                                     Convert.ToString(ri["MachineDescription"]),
                                                     Convert.ToString(ri["FaultPart"]),
                                                     Convert.ToString(ri["FaultMemo"]),
                                                     Convert.ToDateTime(ri["RequestFinishDate"]).ToString("yy/MM/dd HH:mm"),
                                                     Convert.ToString(ri["StatusName"]),
                                                     InspectionDate > Convert.ToDateTime("2000-01-01") ? InspectionDate.ToString("yy/MM/dd HH:mm") : ""
                                                    );



                        }
                        MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", md.Rows.Count));
                        MyRecord.Say("创建SendMail。");
                        MyBase.SendMail sm = new MyBase.SendMail();
                        MyRecord.Say("加载邮件内容。");
                        sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, EDate, brs, NowTime, MyBase.CompanyTitle));
                        sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_设备维修申请单进度追踪统计。", NowTime, MyBase.CompanyTitle));
                        //string mailto = ConfigurationManager.AppSettings["MachineMailTo"], mailcc = ConfigurationManager.AppSettings["MachineMailCC"];
                        //MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        //sm.MailTo = mailto;
                        //sm.MailCC = mailcc;
                        ////sm.MailTo = "my80@my.imedia.com.tw";

                        MyConfig.MailAddress mAddress = MyConfig.GetMailAddress("Machine");
                        MyRecord.Say(string.Format("MailTO:{0}\r\nMailCC:{1}", mAddress.MailTo, mAddress.MailCC));
                        sm.MailTo = mAddress.MailTo;
                        sm.MailCC = mAddress.MailCC;
                        //sm.MailTo = "my80@my.imedia.com.tw";

                        MyRecord.Say("发送邮件。");
                        sm.SendOut();
                        MyRecord.Say("已经发送。");
                    }
                    else
                    {
                        MyRecord.Say("没有找到资料。");
                    }
                }
                MyRecord.Say("------------------发送完成----------------------------");
            }
            catch (Exception e)
            {
                MyRecord.Say(e);
            }
        }

        void MachineRepairReportChecker()
        {
            DateTime rxtime = DateTime.Now;
            string SQL = "";
            MyRecord.Say("---------------------启动核销维修单。------------------------------");
            DateTime NowTime = DateTime.Now;
            DateTime StopTime = NowTime.Date.AddDays(-14);
            MyRecord.Say(string.Format("核销截止时间：{0:yy/MM/dd HH:mm}", StopTime));
            #region 审核纪律单
            try
            {
                MyRecord.Say("核销维修单");
                SQL = @"Select * from _GS_EquipmentRepair a Where a.InputDate < @InputEnd And isNull(a.Status,0)=0  ";
                MyData.MyDataTable mTableEquipmentRepair = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@InputEnd", StopTime, MyData.MyDataParameter.MyDataType.DateTime));
                if (_StopAll) return;
                if (mTableEquipmentRepair != null && mTableEquipmentRepair.MyRows.Count > 0)
                {
                    string memo = string.Format("读取了{0}条记录，下面开始核销....", mTableEquipmentRepair.MyRows.Count);
                    MyRecord.Say(memo);
                    int mTableEquipmentRepairCount = 1;
                    MyData.MyCommand mcd = new MyData.MyCommand();
                    foreach (MyData.MyDataRow r in mTableEquipmentRepair.MyRows)
                    {
                        if (_StopAll) return;
                        int CurID = Convert.ToInt32(r["_ID"]);
                        string RdsNo = Convert.ToString(r["RdsNo"]);
                        memo = string.Format("核销第{0}条，维修单号：{1}，ID：{2}，输入时间：{3:yy/MM/dd HH:mm}", mTableEquipmentRepairCount, RdsNo, CurID, r["InputDate"]);
                        MyRecord.Say(memo);
                        MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
						{
							new MyData.MyDataParameter("@memo","超出15日没有审核，被系统自动核销。"),
							new MyData.MyDataParameter("@id", CurID, MyData.MyDataParameter.MyDataType.Int),
							new MyData.MyDataParameter("@RdsNo",RdsNo),
							new MyData.MyDataParameter("@Status",Convert.ToInt32(r["Status"]),MyData.MyDataParameter.MyDataType.Int)
						};
                        string CheckSQL = "Update [_GS_EquipmentRepair] Set Status =-1,FinishMemo=@memo,FinishTime=GetDate() Where [_ID]=@id";
                        string CheckLogSQL = @"Insert Into [_GS_EquipmentRepair_StatusRecorder](zbid,Author,[state],memo,rdsno,type,typeid,CheckIn,CheckOut)
																						 Values(@id,'系統核銷',0-(@Status+1),@memo,@RdsNo,'核销',2,1,1)";
                        mcd.Add(CheckSQL, string.Format("Update{0}", mTableEquipmentRepairCount), mps);
                        mcd.Add(CheckLogSQL, string.Format("Insert{0}", mTableEquipmentRepairCount), mps);
                        mTableEquipmentRepairCount++;
                    }
                    if (!mcd.Execute())
                    {
                        MyRecord.Say("核销出错。");
                    }
                }
                else
                {
                    MyRecord.Say("没有获取到任何内容");
                }
                MyRecord.Say("核销设备维修单，完成");
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
            #endregion
        }

        #endregion

        #region 计算采购入库到采购单，和计算采购入库上的请购单号

        void PurchaseCalculateLoader()
        {
            MyRecord.Say("开启计算采购单和采购入库单..........");
            Thread t = new Thread(new ThreadStart(PurchaseCalculate));
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启计算采购单和采购入库单成功。");
        }

        void PurchaseCalculate()
        {
            try
            {
                MyData.MyCommand mcd = new MyData.MyCommand();
                string SQL =
    @"
Update b Set b.InStockNumb=ISNULL((Select Sum(Numb) from coPurchStocklst c Where PurchNo=a.rdsno And PID = b.id),0)
  From coPurchase a,coPurchlst b Where a.id=b.zbid And a.status > 0 And b.CloseDate is Null And a.enddate is Null
";
                mcd.Add(SQL, "SQL_1");
                /*
                string SQL1 =
    @"
    Update b Set b.Status=(Case When isNull(b.InStockNumb,0) = 0 Then 1 Else (Case When b.Status > 2 Then b.Status Else 2 End) End),b.CloseDate = (Case When isNull(b.SheetNumb,0) - isNull(b.InStockNumb,0)<=0 Then GetDate() Else Null End) From coPurchase a,coPurchlst b Where a.id=b.zbid And a.status > 0 And a.enddate is Null
    ";
                mcd.Add(SQL1);
                string SQL2 =
    @"
    Update a Set a.Status = (Select Max(Status) From coPurchlst b Where b.zbid=a.id) From coPurchase a Where isNull(a.Status,0)>0 And a.CloseDate is Null And a.enddate is Null
    ";
                mcd.Add(SQL2);
                 */

                string SQL3 = @"
Update a Set a.RequestRdsNo=(Select n.RequestNo From coPurchase m,coPurchLst n Where m.id=n.zbid And m.RdsNo=a.PurchNo And n.id = a.pid)
From coPurchStockLst a Where a.RequestRdsNo is Null";

                mcd.Add(SQL3, "SQL_3");

                string SQL4 = @"

Update a Set a.ReqID=(Select Top 1 n.ReqPID From coPurchase m,coPurchLst n Where m.id=n.zbid And m.RdsNo=a.PurchNo And n.id = a.pid)
From coPurchStockLst a Where a.ReqID is Null
";
                mcd.Add(SQL4, "SQL_4");
                DateTime nowTime = DateTime.Now;
                MyRecord.Say("开始更新。");
                if (mcd.Execute())
                {
                    MyRecord.Say(string.Format("更新成功，耗时：{0}秒", (DateTime.Now - nowTime).TotalSeconds));
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        #endregion

        #region 计算订单送货数

        void OrderCalculateLoader()
        {
            MyRecord.Say("开启计算订单送货数..........");
            Thread t = new Thread(new ThreadStart(OrderCalculate));
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启计算订单送货数成功。");
        }

        void OrderCalculate()
        {
            try
            {
                MyData.MyCommand mcd = new MyData.MyCommand();
                string SQL = @"
Update b Set b.ShippingNumber = isNull((Select Sum(x.SendNumb) From coShiplst x Where x.OrderNo=a.RdsNo And x.Pid=b.id),0)
		From coOrder a,coOrderProd b
 Where a.id=b.zbid And a.id in (Select m.id from coShip m Where m.inputdate between DateAdd(HH,-3,GetDate()) And GetDate())
";
                mcd.Add(SQL, "SQL_1");
                DateTime nowTime = DateTime.Now;
                MyRecord.Say("开始更新。");
                if (mcd.Execute())
                {
                    MyRecord.Say(string.Format("更新成功，耗时：{0}秒", (DateTime.Now - nowTime).TotalSeconds));
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        #endregion

        #region 计算库存的平均周转天数和最后出入库时间，价格，反馈入库日期入库信息到出库

        void StockCalculateLoader()
        {
            MyRecord.Say("开启库存--后台计算..........");
            Thread t = new Thread(new ThreadStart(StockCalculate));
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启库存--后台计算成功。");
        }

        void StockCalculate()
        {
            try
            {
                MyData.MyCommand mcd = new MyData.MyCommand();
                ///计算最后出入库时间
                string SQL =
    @"
Set  NoCount ON

Select Code,LastOutStockDate = Convert(DateTime,Null),
			LastInStockDate = Convert(DateTime,Null),
			StockDate = Convert(DateTime,Null),
			AdjustStockDate = Convert(DateTime,Null),
			StkDate = Convert(DateTime,Null) Into #T From AllMaterialView

Select Code,Max(StockDate) StkDate Into #S From [_WH_OutStockLotNo_View] Group by Code

Select Code,Max(InStockDate) StkDate Into #U from [_WH_InStockLotNo_View] Group by Code

Select a.Code,Min(b.InStockDate) StkDate Into #ST From [_ST_StockListByLotNoView] a Inner Join [_WH_InStockLotNo_View] b ON a.LotNo = b.LotNo Where a.StockNumb>0 Group by a.Code

Select xa.Code,Min(xi.InStockDate) StkDate Into #SA from stAdjustStocklst xa Inner Join [_WH_InStockLotNo_View] xi ON xa.LotNo = xi.LotNo Where xa.Numb >0 Group by xa.Code

Update a Set a.StockDate = s.StkDate From #T a,#ST s Where a.Code= s.Code

Update a Set a.LastOutStockDate = s.StkDate From #T a,#S s Where a.Code= s.Code

Update a Set a.LastInStockDate = s.StkDate From #T a,#U s Where a.Code= s.Code

Update a Set a.AdjustStockDate = s.StkDate From #T a,#SA s Where a.Code= s.Code

Update a Set a.StkDate = Case When a.StockDate is Null Then
							  Case When a.LastOutStockDate is Null Then
								   IsNull(a.LastInStockDate,a.AdjustStockDate)
							  Else
								   a.LastOutStockDate
							  End
						 Else
							  Case When a.LastOutStockDate is Null Then
								   a.StockDate
							  Else
								   Case When a.LastOutStockDate > a.StockDate Then
										a.LastOutStockDate
								   Else
										a.StockDate
								   End
							  End
						 End
		From #T a

Update pbMaterial Set LastOutStockDate = a.StkDate,LotInStockDate = a.StockDate From pbMaterial,#T a Where a.Code=pbMaterial.Code

Update pbKzmsg Set LastOutStockDate = a.StkDate,LotInStockDate = a.StockDate From pbKzmsg,#T a Where a.Code=pbKzmsg.Code

Update pbProduct Set LastOutStockDate = a.StkDate,LotInStockDate = a.StockDate From pbProduct,#T a Where a.Code=pbProduct.Code

--Select * from #T Where StockDate is Not Null

Drop Table #T
Drop Table #ST
Drop Table #S
Drop Table #U
Drop Table #SA

Set NoCount Off
";
                mcd.Add(SQL, "SQL_1");
                ///计算平均库存周转天数
                string SQL3 = @"
Set NoCount ON
Declare @LastDate DateTime

Set @LastDate =(Select Max(SumDate) From stSumStock)

Select b.Code,b.LotNo,a.InStockDate,b.StockDate as OutStockDate,a.numb as InNumb,b.Numb as OutNumb Into #S
  from [_WH_OutStockLotNo_View] b Inner Join [_WH_InStockLotNo_View] a ON b.LotNo=a.LotNo
 Where b.StockDate > @LastDate

Select Code,LotNo,Ceiling(Sum(DateDiff(DD,InStockDate,OutStockDate) * OutNumb) / Sum(OutNumb)) as StkDate,InNumb=Max(InNumb),OutNumb=Sum(OutNumb),StkNumb=Max(InNumb)-Sum(OutNumb),
	   Case When Max(InNumb)-Sum(OutNumb) >0 Then DateDiff(DD,Max(InStockDate),GetDate()) Else 0 End as StkDate2
  Into #T
  From #S Group by Code,LotNo

Select Code,Ceiling(((Ceiling(Sum(StkDate * InNumb) / Sum(InNumb)) * Sum(OutNumb)) +  (Case When Sum(StkNumb)=0 Then 0 Else Ceiling(Sum(StkNumb * StkDate2) / Sum(StkNumb)) End * Sum(StkNumb))) / Sum(InNumb)) as StkDate,Sum(InNumb) as InNumb,Sum(OutNumb) as OutNumb,StkNumb=Sum(StkNumb)
Into #X
From #T Group by Code

Update pbMaterial Set InventoryTurnOverDays = a.StkDate From pbMaterial,#X a Where a.Code=pbMaterial.Code

Update pbKzmsg Set InventoryTurnOverDays = a.StkDate From pbKzmsg,#X a Where a.Code=pbKzmsg.Code

Update pbProduct Set InventoryTurnOverDays = a.StkDate From pbProduct,#X a Where a.Code=pbProduct.Code

Drop Table #S
Drop Table #T
Drop Table #X

Set NoCount Off
";
                mcd.Add(SQL3, "SQL_3");
                ///修改出库的入库时间
                string SQL4 = @"
Set NoCount ON

Update stOutProdlst Set InStockDate = (Select Max(InStockDate) From [_WH_InStockLotNo_View] b Where b.LotNo=stOutProdlst.BatchNo) Where InStockDate is Null

Update stOtherOutlst Set InStockDate = (Select Max(InStockDate) From [_WH_InStockLotNo_View] b Where b.LotNo=stOtherOutlst.BatchNo) Where InStockDate is Null

Update stAdjustStocklst Set InStockDate = (Select Max(InStockDate) From [_WH_InStockLotNo_View] b Where b.LotNo=stAdjustStocklst.BatchNo) Where InStockDate is Null

Update coShiplst Set InStockDate = (Select Max(InStockDate) From [_WH_InStockLotNo_View] b Where b.LotNo=coShiplst.BatchNo) Where InStockDate is Null

Set NoCount Off
";
                mcd.Add(SQL4, "SQL_4");
                DateTime nowTime = DateTime.Now;
                MyRecord.Say("开始库存计算和更新。");
                if (mcd.Execute())
                {
                    MyRecord.Say(string.Format("更新成功，耗时：{0}秒", (DateTime.Now - nowTime).TotalSeconds));
                }
                else
                {
                    MyRecord.Say(string.Format("提交SQL出错，库存计算出错，耗时：{0}秒", (DateTime.Now - nowTime).TotalSeconds));
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        #endregion

        #region 将考勤系统的人员汇入

        void ImportEmployeeLoader()
        {
            MyRecord.Say("开启将考勤系统的数据汇入..........");
            Thread t = new Thread(new ThreadStart(ImportEmployee));
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启将考勤系统的数据汇入成功。");
        }

        void ImportEmployee()
        {
            try
            {
                MyData.MyCommand mcd = new MyData.MyCommand();
                string SQL = @"
Truncate Table _HR_Employee_Register_FromAttendance ";
                string SQLInsert = @"
Insert Into _HR_Employee_Register_FromAttendance(Person_ID,Code,Name,Card_No,Card_SN,UserID)
SELECT a.*,b.[_ID] FROM OPENROWSET('SQLOLEDB','server=192.168.1.10;uid=sa','Select Person_ID,Person_No,Person_Name,Card_No,Card_SN From STCard_ENP.dbo.ST_Person Where Is_Del =0') a Left Outer Join _HR_Employee b ON a.Person_No=b.code
";
                mcd.Add(SQL, "SQL_1");
                mcd.Add(SQLInsert, "SQL_Insert");
                DateTime nowTime = DateTime.Now;
                MyRecord.Say("引入开始");
                if (mcd.Execute())
                {
                    MyRecord.Say(string.Format("考勤数据更新成功，耗时：{0}秒", (DateTime.Now - nowTime).TotalSeconds));
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        #endregion

        #region 定时踢出特殊权限

        void KickOutSpecialHalfHour()
        {
            Thread t = new Thread(KickOutSpecialByHalfHour);
            t.IsBackground = true;
            t.Start();
        }

        void KickOutSpecialByHalfHour()
        {
            MyRecord.Say("---------------------定时踢出特殊权限---------------------------------");
            DateTime bStartTime = DateTime.Now;
            DateTime NowTime = DateTime.Now, TimeSet = DateTime.MinValue;
            try
            {
                string SQL = "";
                SQL = @"UPDATE a SET a.ComputerIP = null,a.ComputerName = null,LogoutTime = GETDATE() FROM _HR_Employee a WHERE a.isSpecial = 1 AND a.LogoutTime < a.LoginTime AND DATEDIFF(MI,a.LoginTime,GETDATE()) > 30 ";
                MyData.MyCommand cmd = new MyData.MyCommand();
                if (_StopAll) return;
                if (cmd.Execute(SQL))
                {
                    MyRecord.Say("执行完毕，运行正常。");
                }
                else
                {
                    MyRecord.Say("执行出错。");
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
            MyRecord.Say("-----------------------定时踢出特殊权限-完毕-------------------------------");
        }
        #endregion

    }
}
