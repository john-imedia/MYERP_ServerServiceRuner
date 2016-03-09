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

namespace MYERP_ServerServiceRuner
{
    public partial class MainService : ServiceBase
    {
        #region 系统处理
        public MainService()
        {
            InitializeComponent();
        }

        public DateTime CheckStockStartTime, CalculatePruchaseStartTime, CalculateOLDPayToNew;
        /// <summary>
        /// 上一次计算完工数领料数计算时间。
        /// </summary>
        public DateTime ProduceFeedBackLastRunTime;

        public bool FirstStart = false;

        protected override void OnStart(string[] args)
        {
            MainTimer.Interval = 1000;
            MainTimer.Elapsed += MainTimer_Elapsed;
            MyRecord.Say(MyBase.ConnectionString);
            ProduceFeedBackLastRunTime = CheckStockStartTime = CalculatePruchaseStartTime = DateTime.Now;
            MyRecord.Say("服务启动");
            FirstStart = true;
            MainTimer.Start();
        }
        #endregion

        void MainTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                DateTime NowTime = DateTime.Now;
                int h = NowTime.Hour, m = NowTime.Minute, s = NowTime.Second, d = NowTime.Day;
                DayOfWeek w = NowTime.DayOfWeek;
                if (FirstStart)
                {
                    FirstStart = false;
                    //ProdPlanForSaveLoader(NowTime);
                    //ProblemSolving8DReportLoder();
                    //SendProduceDiffNumbEmailLoder();
                    //UpdateProduceNoteFnishedNumberLoader();
                    //DateTime xdate = new DateTime(2015, 7, 1);
                    //for (int i = 0; i < 18; i++)
                    //{
                    //    DateTime xxd1 = xdate.AddDays(i).Date.AddHours(11).AddMinutes(55);
                    //    DateTime xxd2 = xdate.AddDays(i).Date.AddHours(23).AddMinutes(55);
                    //    ProdPlanForSave(xxd1);
                    //    ProdPlanForSave(xxd2);
                    //}
                }

                if (h == 0 && m == 0 & s == 0) //计时器归零
                {   ///每天0点对表，计时器归零。
                    ProduceFeedBackLastRunTime = CheckStockStartTime = CalculatePruchaseStartTime = DateTime.Now;
                }

                ///每四个小时候审核一次仓库单据
                if ((NowTime - CheckStockStartTime).TotalHours > 4)
                {
                    CheckStockStartTime = NowTime;
                    if (!CheckStockRecordRunning)
                    {
                        CheckStockRecordRunning = true;
                        CheckStockRecordLoder();   //定時審核出入庫單。
                        Thread.Sleep(500);
                    }
                }
                //每隔2.5个小时计算一次采购数量。
                if ((NowTime - CalculatePruchaseStartTime).TotalHours > 2.5)
                {
                    CalculatePruchaseStartTime = NowTime;
                    PurchaseCalculateLoader();
                    Thread.Sleep(500);
                }

                if ((h == 10 || h == 22) && m == 35 && s == 0) //审核排程，发达成率
                {
                    ConfirmProcessPlan(); //每天10点定时审核单据，先审核单据。
                    //DateTime xdate = new DateTime(2015, 7, 13);
                    //for (int i = 0; i < 18; i++)
                    //{
                    //    DateTime xxd1 = xdate.Date.AddHours(10).AddMinutes(35);
                    //    DateTime xxd2 = xdate.Date.AddHours(22).AddMinutes(35);
                    //    SendProdPlanEmail(xxd1);
                    //    SendProdPlanEmail(xxd2);
                    //}
                    Thread.Sleep(1000);
                    SendProdPlanEmail(NowTime);  //定時發送排程和達成率。
                }
                else if ((h == 11 || h == 23) && m == 55 && s == 0) //保存达成率到月报表，审核纪律单。
                {
                    MyRecord.Say("开启保存达成率线程");
                    ProdPlanForSaveLoader(NowTime);

                }
                else if (h == 6 && m == 17 && s == 12) //发送不良超100%
                {
                    RejectSendMailLoader();  ///发送不良率超过100%的列表，每天早6点发送。
                }
                else if (h == 7 && m == 17 && s == 12) //发送未审核工单
                {
                    if (w == DayOfWeek.Monday || w == DayOfWeek.Tuesday || w == DayOfWeek.Wednesday || w == DayOfWeek.Thursday || w == DayOfWeek.Friday || w == DayOfWeek.Saturday)
                    {
                        SendProduceDayReportLoder();  ///每天发送未审核工单，每天早7点发送。
                        Thread.Sleep(1000);
                        SendProduceUnFinishEmailLoder();  ///未结单工作日发送。
                    }
                }
                else if (h == 0 && m == 5 && s == 7) //发送工单差异数
                {
                    SendProduceDiffNumbEmailLoder();   ///每天零点发送昨天工单差异数
                }
                else if (h == 6 && m == 2 && s == 1) //自动计算库存的最后出库日期，平均周转天数，反馈入库时间到出库表
                {
                    StockCalculateLoader();
                }
                else if ((h == 6 && m == 35 && s == 10)) //清理排程
                {
                    if (!PlanRecordCleanRuning) PlanRecordCleanLoder(); ///每天清理排程
                }
                else if (h == 7 && m == 17 && s == 27)   ///每天发送8D报告跟踪表
                {
                    ProblemSolving8DReportLoder();
                }
                else if ((h == 11 && h == 23) && m == 59 && s == 57) //自动计算完工数
                {
                    ProduceFeedBackLastRunTime = NowTime;
                    if (!ProduceFeedBackRuning)
                    {
                        ProduceFeedBackRuning = true;
                        ProduceFeedBackLoder();
                        Thread.Sleep(500);
                    }
                }
                else if (h == 12 && m == 35 && s == 55)  //自动发送未结束维修申请
                {
                    MachineRepairReportLoder();
                }
                else if (h == 11 && m == 51 && s == 51) //自动计算“三日出货计划异常”异常项目发送邮件。
                {
                    DeliverPlanFinishStatisticErrorSender();
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
                MyData.MyParameter[] mp = new MyData.MyParameter[3] 
                { 
                    new MyData.MyParameter("@zbid", CurrentID, MyData.MyParameter.MyDataType.Int),
                    new MyData.MyParameter("@memo", null),
                    new MyData.MyParameter("@rdsno", CurrentRdsNO)
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
                if ((DateTime.Now - rxtime).TotalMinutes > 20) return;
                Thread.Sleep(5000);
            };
            string SQL = "";
            MyRecord.Say("---------------------启动定时审核出入库单据。------------------------------");
            DateTime NowTime = DateTime.Now;
            DateTime StopTime = NowTime.AddHours(-3.75);
            MyRecord.Say(string.Format("审核截止时间：{0:yy/MM/dd HH:mm}", StopTime));
            CheckStockRecordRunning = true;
            #region 審核產成品入庫----新版
            try
            {
                MyRecord.Say("审核产成品入库单——新ERP系统审核");
                SQL = @"Select * from stProduceStock Where InputDate < @InputEnd And isNull(Status,0)=0 And isono='NEWERP'";
                MyData.MyDataTable mTableProduceInStock = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
                MyData.MyDataTable mTableOtherInStock = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
                MyData.MyDataTable mTableOtherInStock = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
                MyData.MyDataTable mTableProduceInStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
                        if (!mc.Execute(UpdateOLD_SQL, new MyData.MyParameter("@rdsno", RdsNo)))
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
                MyData.MyDataTable mTableOtherInStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
                        if (!mc.Execute(UpdateOLD_SQL, new MyData.MyParameter("@rdsno", RdsNo)))
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
                SQL = @"Select * from coPurchStock Where InputDate < @InputEnd And isNull(Status,0)=0 ";
                MyData.MyDataTable mTablePurchInStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
                        if (!mc.Execute(UpdateOLD_SQL, new MyData.MyParameter("@rdsno", RdsNo)))
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
                MyData.MyDataTable mTableOtherInStock = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
                MyData.MyDataTable mTableProduceOutStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
                        if (!mc.Execute(UpdateOLD_SQL, new MyData.MyParameter("@rdsno", RdsNo)))
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
                MyData.MyDataTable mTableOthereOutStock_OLD = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
                        if (!mc.Execute(UpdateOLD_SQL, new MyData.MyParameter("@rdsno", RdsNo)))
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
                MyData.MyDataTable mTableOtherInStock = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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

            #region 審核库存调整单----新版&旧版
            try
            {
                MyRecord.Say("审核库存调整单——新ERP系统审核");
                SQL = @"Select * from stAdjustStock Where InputDate < @InputEnd And isNull(Status,0)=0";
                MyData.MyDataTable mTableAdjustStock = new MyData.MyDataTable(SQL, new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime));
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
            CheckStockRecordRunning = false;
            MyRecord.Say("-----------------------审核完成。-------------------------------");
        }

        #endregion

        #region 排程达成率发邮件/排程达成率保存

        #region 数据源

        int _PlanExtendTime = 135;

        private void Plan_LoadDataSource(DateTime NowTime, int classtype)
        {
            string SQLPlan = @"
Select a.RdsNo,
       a.Process,
	   a.Department,
       a.PlanBegin,
       a.PlanEnd,
	   a.Inputer,
	   a.Checker,
	   b.MachineCode,
       b.ReqNumb,
	   b.Edd,
	   b.Bdd,
	   b.ReqTime,
	   b.wLevel as WorkLevel,
	   b.capF as CapacityRate,
	   b.PrepF as PrepRate,
	   b.Capacity,
	   b.PrepTime,
	   b.HrNumb,
	   b.PartNO as PartID,
       b.ProdNo,
       b.ProdCode as ProductCode,
       b.Remark,
       b.PN as ColNumb,
	   b.Side,
       b.AutoDuplexPrint,
       MachineName = (Select Name From moMachine Where Code=b.MachineCode),
       FinishMeasurement = (Select FinishMeasurement From moMachine Where Code = b.MachineCode),
       ProcessName = (Select Name from moProcedure where code = a.Process),
       ProductName = (Select Name from pbProduct Where Code= b.ProdCode),
       DepartmentName = (Select Name from pbDept Where [_id]=a.Department),
       FullSortID = (Select FullSortID from pbDept Where [_id]=a.Department),
	   ProdSide=(Case When b.ProcNo='2000' Then (Select Case When (ca1+ca2)>0 And (Cb1+Cb2)>0 Then 2 Else 1 End from moProdProcedure mp Where mp.zbid=b.PRODID And mp.ID=b.ProcID) Else 1 END)
from _PMC_ProdPlan a Inner Join _PMC_ProdPlan_List b ON a.[_ID]=b.zbid 
Where a.[Status] >=2 And isNull(b.YieldOff,0)=0 
  And a.PlanBegin > @DateBegin And a.PlanEnd < @DateEnd And b.Edd< DateAdd(mi,@PlanExtendTime,a.PlanEnd) And b.Bdd > '2000-01-01 00:00:00' 
Order by a.Department,a.Process,b.MachineCode,b.[_id]
";
            string SQLFinish = @"
Select ProcessID as ProcessCode,Numb1,MachinID as MachineCode,StartTime,EndTime,ProduceNo,Inputdate,Inputer,Numb2,Operator,Remark,
       PartID,Remark2,ColNumb,NAColNumb,AdjustNumb,SampleNumb,Rejector,RejectDate,PaperColNumb,ProductCode,AccNoteRdsNo,Side,AutoDuplexPrint
 from prodDailyReport where RptDate Between @fBeginDate And @fEndDate 
";
            DateTime TimeBegin = DateTime.MinValue, TimeEnd = DateTime.MinValue, fTimeBegin = DateTime.MinValue, fTimeEnd = DateTime.MinValue;
            if (classtype == 1) //白班
            {
                fTimeBegin = NowTime.Date.AddHours(7).AddMinutes(55);
                fTimeEnd = NowTime.Date.AddDays(3).AddHours(22);
                TimeBegin = NowTime.Date.AddHours(7).AddMinutes(55);
                TimeEnd = NowTime.Date.AddHours(22).AddMinutes(15);
            }
            else  //夜班
            {
                fTimeBegin = NowTime.Date.AddHours(19).AddMinutes(55);
                fTimeEnd = NowTime.Date.AddDays(4).AddHours(10);
                TimeBegin = NowTime.Date.AddHours(19).AddMinutes(55);
                TimeEnd = NowTime.Date.AddDays(1).AddHours(8).AddMinutes(15);
            }
            MyData.MyParameter[] mfps = new MyData.MyParameter[]
            {
                new MyData.MyParameter("@DateBegin", TimeBegin, MyData.MyParameter.MyDataType.DateTime),
                new MyData.MyParameter("@DateEnd",TimeEnd , MyData.MyParameter.MyDataType.DateTime),
                new MyData.MyParameter("@fBeginDate", fTimeBegin, MyData.MyParameter.MyDataType.DateTime),
                new MyData.MyParameter("@fEndDate", fTimeEnd, MyData.MyParameter.MyDataType.DateTime),
                new MyData.MyParameter("@PlanExtendTime",_PlanExtendTime , MyData.MyParameter.MyDataType.Int)
            };
            DateTime _StartTime = DateTime.Now;
            MyRecord.Say(string.Format("3.1.1排程起：{0}止：{1}，完工单起：{2}止：{3}。", TimeBegin, TimeEnd, fTimeBegin, fTimeEnd));
            MyRecord.Say(string.Format("3.1.2正在读取生产计划。AT:{0}", DateTime.Now));
            MyData.MyDataTable md1 = new MyData.MyDataTable(SQLPlan, 100, mfps);
            var v1 = from a in md1.MyRows select new Plan_Item(a);
            MyRecord.Say(string.Format("3.1.3生产计划已经获得，正在从服务器读取到本机。耗时：{0}", (DateTime.Now - _StartTime).TotalMinutes));
            Plan_Lists = v1.ToList();
            MyRecord.Say(string.Format("3.1.4正在读取生产完工单。AT:{0}", DateTime.Now));
            _StartTime = DateTime.Now;
            MyData.MyDataTable md2 = new MyData.MyDataTable(SQLFinish, 100, mfps);
            var v2 = from a in md2.MyRows select new Plan_FinishItem(a);
            MyRecord.Say(string.Format("3.1.5完工单已经获得，正在从服务器读取到本机。耗时：{0}", (DateTime.Now - _StartTime).TotalMinutes));
            Plan_FinishLists = v2.ToList();
            MyRecord.Say("3.1.6数据源读取完毕");
        }

        #endregion

        #region 计算时使用的类。

        List<Plan_Item> Plan_Lists = new List<Plan_Item>();
        List<Plan_FinishItem> Plan_FinishLists = new List<Plan_FinishItem>();

        class Plan_GridItem
        {
            public Plan_GridItem(IGrouping<object, Plan_Item> g)
            {
                if (g.Count() > 0)
                {
                    Plan_Item curItem = g.FirstOrDefault();
                    RdsNo = curItem.RdsNo;
                    DepartmentID = curItem.DepartmentID;
                    ProcessCode = curItem.ProcessCode;
                    ProcessName = curItem.ProcessName;
                    MachineCode = curItem.MachineCode;
                    MachineName = curItem.MachineName;
                    DepartmentFullSortID = curItem.DepartmentFullSortID;
                    DepartmentName = curItem.DepartmentName;

                    hrNumb = g.Max(x => x.HrNumb);
                    Inputer = curItem.Inputer;
                    Checker = curItem.Checker;
                    var vd = from a in g
                             where (a.Bdd - a.PlanBegin).TotalMinutes > -5 && (a.Edd - a.PlanEnd).TotalMinutes < 5
                             select a;
                    if (vd != null && vd.Count() > 0)
                    {
                        BDD = vd.Min(x => x.Bdd);
                        EDD = vd.Max(x => x.Edd);
                        if (curItem.ProcessCode == "2000" || curItem.FinishMeasurement == 1)
                        {
                            var vf = from a in g
                                     where a.FinishSheetNumb > 0
                                     select a;
                            PlanCount = vd.Count();
                            FinishCount = vf.Count();

                            PlanSheetNumb = vd.Sum(x => x.ReqNumb);
                            PlanProdNumb = vd.Sum(x => x.ReqNumb); // vd.Sum(m => m.ReqNumb * m.ColNumb);
                            if (vf != null && vf.Count() > 0)
                            {
                                FinishSheetNumb = vf.Sum(x => x.FinishSheetNumb);
                                FinishProdNumb = vf.Sum(x => x.FinishSheetNumb); //vf.Sum(m => m.FinishProdNumb);
                            }
                        }
                        else
                        {
                            var vf = from a in g
                                     where a.Side == 0 && a.FinishSheetNumb > 0
                                     select a;
                            PlanCount = vd.Count();
                            FinishCount = vf.Count();

                            PlanSheetNumb = vd.Sum(x => x.ReqNumb);
                            PlanProdNumb = vd.Sum(m => m.ReqNumb * m.ColNumb);
                            if (vf != null && vf.Count() > 0)
                            {
                                FinishSheetNumb = vf.Sum(x => x.FinishSheetNumb);
                                FinishProdNumb = vf.Sum(m => m.FinishProdNumb);
                            }
                        }
                    }
                }
            }

            public string RdsNo { get; set; }
            public int DepartmentID { get; set; }
            public string DepartmentName { get; set; }
            public string ProcessCode { get; set; }
            public string ProcessName { get; set; }
            public string MachineCode { get; set; }
            public string MachineName { get; set; }
            public double hrNumb { get; set; }
            public string Inputer { get; set; }
            public string Checker { get; set; }
            public DateTime BDD { get; set; }
            public DateTime EDD { get; set; }
            public int PlanCount { get; set; }
            public int FinishCount { get; set; }
            public double PlanSheetNumb { get; set; }
            public double FinishSheetNumb { get; set; }
            public double PlanProdNumb { get; set; }
            public double FinishProdNumb { get; set; }
            public string DepartmentFullSortID { get; set; }

        }

        class Plan_GroupKey : IComparable
        {
            public Plan_GroupKey(string rdsno, int deptID, string PCode, string MCode)
            {
                RdsNo = rdsno;
                DepartmentID = deptID;
                ProcessCode = PCode;
                MachineCode = MCode;
            }

            public string RdsNo { get; set; }
            public int DepartmentID { get; set; }
            public string ProcessCode { get; set; }
            public string MachineCode { get; set; }

            public int CompareTo(Object obj)
            {
                Plan_GroupKey other = obj as Plan_GroupKey;
                int result = RdsNo.CompareTo(other.RdsNo);
                if (result == 0)
                {
                    result = DepartmentID.CompareTo(other.DepartmentID);
                    if (result == 0)
                    {
                        result = DepartmentID.CompareTo(other.DepartmentID);
                        if (result == 0)
                        {
                            result = ProcessCode.CompareTo(other.ProcessCode);
                            if (result == 0)
                            {
                                result = MachineCode.CompareTo(other.MachineCode);
                            }
                        }
                    }
                }
                return result;
            }
            public override string ToString()
            {
                return string.Format("{0},{1},{2},{3}", RdsNo, DepartmentID, ProcessCode, MachineCode);
            }
        }

        class Plan_Item
        {
            /// <summary>
            ///  加载工单计划类
            /// </summary>
            public Plan_Item(MyData.MyDataRow r)
            {
                RdsNo = Convert.ToString(r["RdsNo"]);
                ProcessCode = Convert.ToString(r["Process"]);
                DepartmentID = Convert.ToInt32(r["Department"]);
                PlanBegin = Convert.ToDateTime(r["PlanBegin"]);
                PlanEnd = Convert.ToDateTime(r["PlanEnd"]);
                Inputer = Convert.ToString(r["Inputer"]);
                Checker = Convert.ToString(r["Checker"]);
                MachineCode = Convert.ToString(r["MachineCode"]);
                ReqNumb = Convert.ToDouble(r["ReqNumb"]);
                Edd = Convert.ToDateTime(r["Edd"]);
                Bdd = Convert.ToDateTime(r["Bdd"]);
                ReqTime = Convert.ToDouble(r["ReqTime"]);
                WorkLevel = Convert.ToDouble(r["WorkLevel"]);
                CapacityRate = Convert.ToDouble(r["CapacityRate"]);
                PrepRate = Convert.ToDouble(r["PrepRate"]);
                Capacity = Convert.ToDouble(r["Capacity"]);
                PrepTime = Convert.ToDouble(r["PrepTime"]);
                HrNumb = Convert.ToDouble(r["HrNumb"]);
                PartID = Convert.ToString(r["PartID"]);
                ProduceRdsNo = Convert.ToString(r["ProdNo"]);
                ProductCode = Convert.ToString(r["ProductCode"]);
                Remark = Convert.ToString(r["Remark"]);
                ColNumb = Convert.ToInt32(r["ColNumb"]);
                ProdSide = Convert.ToInt32(r["ProdSide"]);
                Side = Convert.ToInt32(r["Side"]);
                AutoDuplexPrint = Convert.ToBoolean(r["AutoDuplexPrint"]);
                DepartmentName = Convert.ToString(r["DepartmentName"]);
                DepartmentFullSortID = Convert.ToString(r["FullSortID"]);
                ProcessName = Convert.ToString(r["ProcessName"]);
                MachineName = Convert.ToString(r["MachineName"]);
            }
            public string RdsNo { get; set; }
            public string ProduceRdsNo { get; set; }
            public string ProductCode { get; set; }
            public string ProcessCode { get; set; }
            public string ProcessName { get; set; }
            public int DepartmentID { get; set; }
            public string DepartmentName { get; set; }
            public string DepartmentFullSortID { get; set; }
            public DateTime PlanBegin { get; set; }
            public DateTime PlanEnd { get; set; }
            public string Inputer { get; set; }
            public string Checker { get; set; }
            public string MachineCode { get; set; }
            public string MachineName { get; set; }
            public double ReqNumb { get; set; }
            public DateTime Edd { get; set; }
            public DateTime Bdd { get; set; }
            public double ReqTime { get; set; }
            public double WorkLevel { get; set; }
            public double CapacityRate { get; set; }
            public double PrepRate { get; set; }
            public double Capacity { get; set; }
            public double PrepTime { get; set; }
            public double HrNumb { get; set; }
            public string PartID { get; set; }
            public string Remark { get; set; }
            public int ColNumb { get; set; }
            public int Side { get; set; }
            public int ProdSide { get; set; }
            public double FinishSheetNumb { get; set; }
            public double FinishProdNumb { get; set; }
            public bool AutoDuplexPrint { get; set; }
            public string Type { get; set; }
            public int FinishMeasurement { get; set; }
        }

        class Plan_FinishItem
        {
            public Plan_FinishItem(MyData.MyDataRow r)
            {
                ProcessCode = Convert.ToString(r["ProcessCode"]);
                FinishNumb = Convert.ToInt32(r["Numb1"]);
                MachineCode = Convert.ToString(r["MachineCode"]);
                StartTime = Convert.ToDateTime(r["StartTime"]);
                EndTime = Convert.ToDateTime(r["EndTime"]);
                ProduceRdsNo = Convert.ToString(r["ProduceNo"]);
                InputDate = Convert.ToDateTime(r["Inputdate"]);
                Inputer = Convert.ToString(r["Inputer"]);
                RejectNumb = Convert.ToInt32(r["Numb2"]);
                OP = Convert.ToString(r["Operator"]);
                Remark = Convert.ToString(r["Remark"]);
                PartID = Convert.ToString(r["PartID"]);
                HandNo = Convert.ToString(r["Remark2"]);
                FinishColNumb = Convert.ToInt32(r["ColNumb"]);
                RejectColNumb = Convert.ToInt32(r["NAColNumb"]);
                AdjustNumb = Convert.ToInt32(r["AdjustNumb"]);
                SampleNumb = Convert.ToInt32(r["SampleNumb"]);
                Rejector = Convert.ToString(r["Rejector"]);
                RejectDate = Convert.ToDateTime(r["RejectDate"]);
                PaperColNumb = Convert.ToInt32(r["PaperColNumb"]);
                ProductCode = Convert.ToString(r["ProductCode"]);
                AccNoteRdsNo = Convert.ToString(r["AccNoteRdsNo"]);
                AutoDuplexPrint = Convert.ToBoolean(r["AutoDuplexPrint"]);
                Side = Convert.ToInt32(r["Side"]);
            }

            public string ProductCode { get; set; }

            public string ProcessCode { get; set; }

            public string MachineCode { get; set; }

            public int FinishNumb { get; set; }

            public int FinishProdNumb
            {
                get
                {
                    return FinishColNumb * FinishNumb;
                }
            }

            public DateTime StartTime { get; set; }
            public DateTime EndTime { get; set; }

            public string ProduceRdsNo { get; set; }

            public DateTime InputDate { get; set; }
            public string Inputer { get; set; }

            public int RejectNumb { get; set; }

            public int RejectProdNumb
            {
                get
                {
                    return RejectColNumb * RejectNumb;
                }
            }

            public string OP { get; set; }
            public string Remark { get; set; }
            public string PartID { get; set; }
            public string HandNo { get; set; }
            public int FinishColNumb { get; set; }
            public int RejectColNumb { get; set; }
            public int AdjustNumb { get; set; }
            public int SampleNumb { get; set; }
            public string Rejector { get; set; }
            public DateTime RejectDate { get; set; }
            public int PaperColNumb { get; set; }
            public string AccNoteRdsNo { get; set; }
            public bool iCalItem { get; set; }
            public int Side { get; set; }
            public bool AutoDuplexPrint { get; set; }
        }

        /// <summary>
        /// 用于保存加载时的顺序分配。
        /// </summary>
        class ProcessSumItemBkp
        {
            /// <summary>
            /// 排程单号
            /// </summary>
            public string RdsNo { get; set; }
            /// <summary>
            /// 工单号
            /// </summary>
            public string ProduceRdsNo { get; set; }
            public string ProductCode { get; set; }
            public string ProcessCode { get; set; }
            public string PartID { get; set; }
            public int Side { get; set; }
            public int DepartmentID { get; set; }
            public string MachineCode { get; set; }
            public double FinishSheetNumb { get; set; }
            public double FinishProdNumb { get; set; }
            public int SumTime { get; set; }
        }

        #endregion

        #region 计算达成率
        private List<Plan_GridItem> Plan_LoadFinishedRate(DateTime NowTime, int classtype)
        {
            MyRecord.Say("3.1 读取资料数据源,并保存到本机...");
            Plan_LoadDataSource(NowTime, classtype);
            MyRecord.Say("3.2 计算达成率和加载表格...");
            MyRecord.Say("    设定各个条件。");
            DateTime fTimeBegin = DateTime.MinValue, fTimeEnd = DateTime.MinValue;
            if (classtype == 1) //白班
            {
                fTimeBegin = NowTime.Date.AddHours(7).AddMinutes(55);
                fTimeEnd = NowTime.Date.AddDays(2).AddHours(22);
            }
            else  //夜班
            {
                fTimeBegin = NowTime.Date.AddHours(19).AddMinutes(55);
                fTimeEnd = NowTime.Date.AddDays(3).AddHours(10);
            }
            List<Plan_GridItem> _GridData = new List<Plan_GridItem>();
            MyRecord.Say(string.Format("3.2  完工单限定时间：fTimeBegin={0}，fTimeEnd={1}", fTimeBegin, fTimeEnd));
            #region 印刷达成率
            MyRecord.Say("3.3 计算印刷排程达成率。");
            var vPrintProduceSource = from a in Plan_Lists
                                      where (a.Bdd - a.PlanBegin).TotalMinutes > -5 && a.ProcessCode == "2000"
                                         && (!(a.AutoDuplexPrint && a.ProdSide == 2 && a.Side == 1))
                                      orderby a.DepartmentID, a.ProcessCode, a.MachineCode
                                      select a;
            List<Plan_Item> _LocalPringSource = vPrintProduceSource.ToList();
            List<ProcessSumItemBkp> _FinishNumbBkp = new List<ProcessSumItemBkp>();
            MyRecord.Say(string.Format("3.3.1 数据源已经获取，逐笔工单计算印刷排程完工数，共{0}笔。", _LocalPringSource.Count));
            foreach (var item in _LocalPringSource)
            {
                var vpmt = from a in _LocalPringSource
                           where a.RdsNo == item.RdsNo && a.MachineCode == item.MachineCode
                           select a.Bdd;
                DateTime StartTime = vpmt.Min();
                var vf = from a in Plan_FinishLists
                         where a.InputDate >= fTimeBegin && a.InputDate <= fTimeEnd && a.ProcessCode == "2000" && a.MachineCode == item.MachineCode
                             && a.StartTime >= StartTime.AddMinutes(-1) && a.EndTime <= item.PlanEnd.AddMinutes(5)
                             && a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID && a.Side == item.Side
                         select a;
                double finishProdNumb = vf.Sum(x => x.FinishProdNumb), finishNumb = vf.Sum(x => x.FinishNumb);
                var vppt = from a in _LocalPringSource
                           where a.RdsNo == item.RdsNo && a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID && a.Side == item.Side
                           select a;
                if (vppt != null && vppt.Count() > 1)
                {
                    var vbkp = from a in _FinishNumbBkp
                               where a.RdsNo == item.RdsNo && a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID && a.MachineCode == item.MachineCode && a.Side == item.Side
                               select a;
                    if (vbkp != null && vbkp.Count() > 0)
                    {
                        ProcessSumItemBkp psp = vbkp.FirstOrDefault();
                        if ((finishProdNumb - psp.FinishProdNumb) <= (item.ReqNumb * item.ColNumb) || psp.SumTime >= (vppt.Count() - 1))
                        {
                            finishProdNumb = finishProdNumb - psp.FinishProdNumb;
                            finishNumb = finishNumb - psp.FinishSheetNumb;
                        }
                        else
                        {
                            finishProdNumb = item.ReqNumb * item.ColNumb;
                            finishNumb = item.ReqNumb;
                        }
                        psp.FinishProdNumb += finishProdNumb;
                        psp.FinishSheetNumb += finishNumb;
                        psp.SumTime += 1;
                        _FinishNumbBkp.Add(psp);
                    }
                    else
                    {
                        ProcessSumItemBkp psp = new ProcessSumItemBkp();
                        if (finishProdNumb > item.ReqNumb * item.ColNumb)
                        {
                            finishProdNumb = item.ReqNumb * item.ColNumb;
                            finishNumb = item.ReqNumb;
                        }
                        psp.RdsNo = item.RdsNo;
                        psp.ProduceRdsNo = item.ProduceRdsNo;
                        psp.PartID = item.PartID;
                        psp.MachineCode = item.MachineCode;
                        psp.Side = item.Side;
                        psp.FinishProdNumb = finishProdNumb;
                        psp.FinishSheetNumb = finishNumb;
                        psp.SumTime = 1;
                        _FinishNumbBkp.Add(psp);
                    }
                }
                if (finishProdNumb <= 0) finishProdNumb = 0;
                if (finishNumb <= 0) finishNumb = 0;
                if (finishNumb > (item.ReqNumb * 1.10) && item.ReqNumb <= 500)   //印刷超数量，要按照需求数计算，不可以超出110%
                {
                    item.FinishProdNumb = item.ReqNumb * item.ColNumb * 1.10;
                    item.FinishSheetNumb = item.ReqNumb * 1.10;
                }
                else
                {
                    item.FinishProdNumb = finishProdNumb;
                    item.FinishSheetNumb = finishNumb;
                }
                item.Type = ((item.Bdd - item.PlanBegin).TotalMinutes > -5 && (item.Edd - item.PlanEnd).TotalMinutes < 5) ? MyConvert.ZHLC("正常") : MyConvert.ZHLC("超出");
                if (finishProdNumb > 0) foreach (var iv in vf) iv.iCalItem = true;
            }
            MyRecord.Say("3.3.2 完工数已经计算，计算印刷达成率");
            var vPrintGrid = from a in _LocalPringSource
                             group a by new { a.RdsNo, a.DepartmentID, a.ProcessCode, a.MachineCode } into g
                             orderby g.Key.MachineCode, g.Key.RdsNo
                             select new Plan_GridItem(g);
            _GridData = vPrintGrid.ToList();
            MyRecord.Say("3.3.3 印刷达成率完成。");
            #endregion
            #region 后制达成率
            MyRecord.Say("3.4 计算后制排程达成率。");
            var vProduceSource = from a in Plan_Lists
                                 where a.Side == 0 && (a.Bdd - a.PlanBegin).TotalMinutes > -5 && a.ProcessCode != "2000"
                                 select a;
            _FinishNumbBkp = null;
            _FinishNumbBkp = new List<ProcessSumItemBkp>();
            MyRecord.Say(string.Format("3.4.1 数据源已经获取，逐笔工单计算后制排程完工数，共{0}笔。", vProduceSource.Count()));
            foreach (var item in vProduceSource)
            {
                var vpmt = from a in Plan_Lists
                           where a.RdsNo == item.RdsNo && a.MachineCode == item.MachineCode
                           select a.Bdd;
                DateTime StartTime = vpmt.Min();
                var vf = from a in Plan_FinishLists
                         where a.InputDate >= fTimeBegin && a.InputDate <= fTimeEnd && a.ProcessCode == item.ProcessCode && a.MachineCode == item.MachineCode
                               && a.StartTime >= StartTime.AddMinutes(-1) && a.EndTime <= item.PlanEnd.AddMinutes(15)
                               && a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID
                         select a;
                double finishProdNumb = vf.Sum(x => x.FinishProdNumb), finishNumb = vf.Sum(x => x.FinishNumb);
                var vppt = from a in vProduceSource
                           where a.RdsNo == item.RdsNo && a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID && a.Side == item.Side && a.MachineCode == item.MachineCode
                           select a;
                if (vppt != null && vppt.Count() > 1)
                {
                    var vbkp = from a in _FinishNumbBkp
                               where a.RdsNo == item.RdsNo && a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID && a.MachineCode == item.MachineCode
                               select a;
                    if (vbkp != null && vbkp.Count() > 0)
                    {
                        ProcessSumItemBkp psp = vbkp.FirstOrDefault();
                        if (((finishProdNumb - psp.FinishProdNumb) <= (item.ReqNumb * item.ColNumb)) || (psp.SumTime >= (vppt.Count() - 1)))
                        {
                            finishProdNumb = finishProdNumb - psp.FinishProdNumb;
                            finishNumb = finishNumb - psp.FinishSheetNumb;
                        }
                        else
                        {
                            finishProdNumb = item.ReqNumb * item.ColNumb;
                            finishNumb = item.ReqNumb;
                        }
                        psp.FinishProdNumb += finishProdNumb;
                        psp.FinishSheetNumb += finishNumb;
                        psp.SumTime += 1;
                        _FinishNumbBkp.Add(psp);
                    }
                    else
                    {
                        ProcessSumItemBkp psp = new ProcessSumItemBkp();
                        if (finishProdNumb > item.ReqNumb * item.ColNumb)
                        {
                            finishProdNumb = item.ReqNumb * item.ColNumb;
                            finishNumb = item.ReqNumb;
                        }
                        psp.RdsNo = item.RdsNo;
                        psp.ProduceRdsNo = item.ProduceRdsNo;
                        psp.PartID = item.PartID;
                        psp.MachineCode = item.MachineCode;
                        psp.FinishProdNumb = finishProdNumb;
                        psp.FinishSheetNumb = finishNumb;
                        psp.SumTime = 1;
                        _FinishNumbBkp.Add(psp);
                    }
                }
                if (finishProdNumb <= 0) finishProdNumb = 0;
                if (finishNumb <= 0) finishNumb = 0;
                item.FinishProdNumb = finishProdNumb;
                item.FinishSheetNumb = finishNumb;
                item.Type = ((item.Bdd - item.PlanBegin).TotalMinutes > -5 && (item.Edd - item.PlanEnd).TotalMinutes < 5) ? MyConvert.ZHLC("正常") : MyConvert.ZHLC("超出");
                if (finishProdNumb > 0) foreach (var iv in vf) iv.iCalItem = true;
            }
            MyRecord.Say("3.4.2 完工数计算完毕，计算后制达成率。");
            var vProduceGrid = from a in vProduceSource
                               group a by new { a.RdsNo, a.DepartmentID, a.ProcessCode, a.MachineCode } into g
                               select new Plan_GridItem(g);
            if (_GridData.Count > 0)
                _GridData.AddRange(vProduceGrid);
            else
                _GridData = vProduceGrid.ToList();
            MyRecord.Say("3.4.3 后制达成率完成。");
            #endregion
            return _GridData;


        }
        #endregion

        void SendProdPlanEmail(DateTime NowTime)
        {
            try
            {
                MyRecord.Say("-------------------------------------------开始计算排程达成率-------------------------------------------");
                MyRecord.Say("1.加载邮件体。");
                #region 表头
                string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{4}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {0:yy/MM/dd} {1} 所有排程及达成率。</FONT></DIV>
<DIV><FONT size=2 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; （说明：完工数只计算良品数；只计算计划开始时间至{3:MM/dd HH时}之间的完工单。）</FONT></DIV>
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
    生产计划单号
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    计划开始时间~结束时间
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    创建人/审核人
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    排期数（产品数）
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    完工数（产品数）
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    产量达成率
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    排单笔数
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    完成笔数
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    笔数达成率
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
                #endregion
                MyRecord.Say("2.加载表格行。");
                #region 表格行
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
    </TR>
";
                #endregion
                int classtype = 1;
                string byb = "";
                if (NowTime.Hour == 10) //10点发送昨日夜班排程
                {
                    classtype = 2;
                    NowTime = NowTime.AddDays(-1);
                    byb = "夜班";
                }
                else if (NowTime.Hour == 22) //22点发送当日白天排程
                {
                    classtype = 1;
                    byb = "白班";
                }
                MyRecord.Say(string.Format("3.处理条件，NowTime={0}，班次：{1}，班次ID：{2}", NowTime, byb, classtype));
                string xLine = string.Empty; int iRow = 0; DateTime t1 = DateTime.Now;
                if (NowTime > DateTime.MinValue && classtype > 0)
                {
                    MyRecord.Say("3.获取并计算达成率");
                    List<Plan_GridItem> _GridData = (Plan_LoadFinishedRate(NowTime, classtype));
                    var vGridDataSource = from a in _GridData
                                          where a.BDD > DateTime.Parse("2000-01-01") && a.EDD > DateTime.Parse("2000-01-01") && a.PlanCount > 0
                                          orderby a.DepartmentFullSortID, a.ProcessCode, a.MachineCode, a.RdsNo
                                          select a;
                    MyRecord.Say("4.达成率计算完毕，开始生成邮件内容。");
                    foreach (var item in vGridDataSource)
                    {
                        double y1 = item.PlanCount != 0 ? Convert.ToDouble(item.FinishCount) / Convert.ToDouble(item.PlanCount) : 0;
                        double y2 = item.FinishProdNumb != 0 ? Convert.ToDouble(item.FinishProdNumb) / Convert.ToDouble(item.PlanProdNumb) : 0;
                        string xbr = string.Format(br,
                            item.DepartmentName,
                            item.ProcessName,
                            item.MachineName,
                            item.RdsNo,
                            string.Format("{0:MM/dd HH:mm}~{1:MM/dd HH:mm}", item.BDD, item.EDD),
                            string.Format("{0}/{1}", item.Inputer, item.Checker),
                            item.PlanProdNumb,
                            item.FinishProdNumb,
                            string.Format("{0:0.00%}", y2),
                            item.PlanCount,
                            item.FinishCount,
                            string.Format("{0:0.00%}", y1));
                        MyRecord.Say(string.Format("4.1 - 第{0}行，部门：{1}，工序：{2}，机台：{3}，笔数达成率：{4}，产量达成率：{5}。", iRow, item.DepartmentName, item.ProcessName, item.MachineName, y1, y2));
                        iRow++;
                        xLine += xbr;
                    }
                    MyRecord.Say(string.Format("表格一共：{0}行，已经生成。总耗时：{1}秒。", iRow, (DateTime.Now - t1).TotalSeconds));
                    MyRecord.Say("创建SendMail。");
                    MyBase.SendMail sm = new MyBase.SendMail();
                    MyRecord.Say("加载邮件内容。");
                    sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, NowTime, byb, xLine, DateTime.Now, MyBase.CompanyTitle));
                    sm.Subject = MyConvert.ZH_TW(string.Format("{2}{0:yy年MM月dd日}{1}排程及达成率", NowTime, byb, MyBase.CompanyTitle));
                    string MailTo = ConfigurationManager.AppSettings["PlanMailTo"], MailCC = ConfigurationManager.AppSettings["PlanMailCC"];
                    MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", MailTo, MailCC));
                    sm.MailTo = MailTo;
                    sm.MailCC = MailCC;
                    MyRecord.Say("发送邮件。");
                    sm.SendOut();
                    MyRecord.Say("已经发送。");
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
            MyRecord.Say("-------------------------------------------计算排程达成率完成-------------------------------------------");
        }


        Thread tProdPlanForSaveing;

        bool _StopProdPlanForSaved = false;

        void ProdPlanForSaveLoader(DateTime NowTime)
        {
            if (tProdPlanForSaveing != null && tProdPlanForSaveing.IsAlive)
            {
                _StopProdPlanForSaved = true;
                DateTime xTime = DateTime.Now;
                do
                {
                    _StopProdPlanForSaved = true;
                } while (tProdPlanForSaveing.IsAlive && (DateTime.Now - xTime).TotalSeconds > 10);
            }
            _StopProdPlanForSaved = false;
            tProdPlanForSaveing = new Thread(new ParameterizedThreadStart(ProdPlanForSaveRunner));
            tProdPlanForSaveing.IsBackground = true;
            tProdPlanForSaveing.Start(NowTime);
        }

        void ProdPlanForSaveRunner(object NowTime)
        {
            DateTime xDate = ((DateTime)NowTime).AddDays(-3).Date;
            MyRecord.Say(string.Format("当前时间：{0}，计算开始时间：{1}", NowTime, xDate));
            for (int i = 0; i < 3; i++)
            {
                if (_StopProdPlanForSaved) return;
                DateTime iDate = xDate.AddDays(i).Date;
                MyRecord.Say(string.Format("计算时间：{0}，白班", iDate));
                ProdPlanForSave(iDate, 1);
                if (_StopProdPlanForSaved) return;
                MyRecord.Say(string.Format("计算时间：{0}，夜班", iDate));
                ProdPlanForSave(iDate, 2);
            }
        }

        void ProdPlanForSave(DateTime NowTime, int classtype)
        {
            try
            {
                MyRecord.Say("-----------开始计算排程达成率For保存------------");
                string byb = "";
                MyData.MyCommand xmcd = new MyData.MyCommand();
                //if (NowTime.Hour == 11) //12点保存3日前夜班排程
                //{
                //    classtype = 2;
                //    NowTime = NowTime.AddDays(-3);
                //    byb = "夜班";
                //}
                //else if (NowTime.Hour == 23) //23点50保存2日前白天排程
                //{
                //    NowTime = NowTime.AddDays(-2);
                //    classtype = 1;
                //    byb = "白班";
                //}
                if (classtype == 1) byb = "白班";
                else if (classtype == 2) byb = "夜班";

                MyRecord.Say(string.Format("1.处理条件，NowTime={0}，班次：{1}，班次ID：{2}，开始计算数据源。", NowTime, byb, classtype));
                if (NowTime > DateTime.MinValue && classtype > 0)
                {
                    MyRecord.Say("3.获取并计算达成率");
                    List<Plan_GridItem> _GridData = (Plan_LoadFinishedRate(NowTime, classtype));
                    var vGridDataSource = from a in _GridData
                                          where a.BDD > DateTime.Parse("2000-01-01") && a.EDD > DateTime.Parse("2000-01-01") && a.PlanCount > 0
                                          orderby a.DepartmentFullSortID, a.ProcessCode, a.MachineCode, a.RdsNo
                                          select a;
                    MyRecord.Say("4.达成率计算完毕，生成保存语句。");
                    string xLine = string.Empty; int iRow = 0; DateTime t1 = DateTime.Now;
                    foreach (var item in vGridDataSource)
                    {
                        if (_StopProdPlanForSaved) return;
                        double y1 = item.PlanCount != 0 ? Convert.ToDouble(item.FinishCount) / Convert.ToDouble(item.PlanCount) : 0;
                        double y2 = item.FinishProdNumb != 0 ? Convert.ToDouble(item.FinishProdNumb) / Convert.ToDouble(item.PlanProdNumb) : 0;
                        string SQLSave = @"
Delete from [_PMC_ProdPlan_YieldRate] Where PlanRdsNo=@PlanRdsNo And MachineCode=@MachineCode
Insert Into [_PMC_ProdPlan_YieldRate](PlanRdsNo,PlanType,PlanBegin,PlanEnd,Inputer,Checker,DeaprtmentID,DepartmentHrNumb,ProcessCode,MachineCode,PlanSheetNumb,FinishSheetNumb,PlanNumb,FinishNumb,PlanProdNumb,FinishProdNumb)
Values(@PlanRdsNo,@PlanType,@PlanBegin,@PlanEnd,@Inputer,@Checker,@DeaprtmentID,@DepartmentHrNumb,@ProcessCode,@MachineCode,@PlanSheetNumb,@FinishSheetNumb,@PlanNumb,@FinishNumb,@PlanProdNumb,@FinishProdNumb)";
                        MyData.MyParameter[] amps = new MyData.MyParameter[]
                        {
                            new MyData.MyParameter("@PlanRdsNo",item.RdsNo),
                            new MyData.MyParameter("@PlanType",classtype, MyData.MyParameter.MyDataType.Int ),
                            new MyData.MyParameter("@PlanBegin",item.BDD, MyData.MyParameter.MyDataType.DateTime),
                            new MyData.MyParameter("@PlanEnd",item.EDD, MyData.MyParameter.MyDataType.DateTime),
                            new MyData.MyParameter("@Inputer",item.Inputer),
                            new MyData.MyParameter("@Checker",item.Checker),
                            new MyData.MyParameter("@DeaprtmentID",item.DepartmentID, MyData.MyParameter.MyDataType.Int),
                            new MyData.MyParameter("@DepartmentHrNumb",item.hrNumb, MyData.MyParameter.MyDataType.Numeric),
                            new MyData.MyParameter("@ProcessCode",item.ProcessCode),
                            new MyData.MyParameter("@MachineCode",item.MachineCode),
                            new MyData.MyParameter("@PlanSheetNumb",item.PlanCount , MyData.MyParameter.MyDataType.Numeric),
                            new MyData.MyParameter("@FinishSheetNumb",item.FinishCount , MyData.MyParameter.MyDataType.Numeric),
                            new MyData.MyParameter("@PlanNumb",item.PlanSheetNumb , MyData.MyParameter.MyDataType.Numeric),
                            new MyData.MyParameter("@FinishNumb",item.FinishSheetNumb , MyData.MyParameter.MyDataType.Numeric),
                            new MyData.MyParameter("@PlanProdNumb",item.PlanProdNumb, MyData.MyParameter.MyDataType.Numeric),
                            new MyData.MyParameter("@FinishProdNumb",item.FinishProdNumb, MyData.MyParameter.MyDataType.Numeric)
                        };
                        xmcd.Add(SQLSave, string.Format("X_Insert{0}", iRow), amps);
                        MyRecord.Say(string.Format("4.3-第{0}行，部门：{1}，工序：{2}，机台：{3}", iRow, item.DepartmentName, item.ProcessName, item.MachineName));
                        iRow++;
                    }

                    MyRecord.Say(string.Format("一共：{0}行，已经生成。总耗时：{1}秒。", iRow, (DateTime.Now - t1).TotalSeconds));
                    MyRecord.Say("开始提交");
                    if (_StopProdPlanForSaved) return;
                    if (xmcd.Execute())
                        MyRecord.Say("保存完毕！");
                    else
                        MyRecord.Say("运行出错。");
                }
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
            MyRecord.Say("-----------保存排程达成率完成-----------");
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
                MyData.MyParameter[] myps = new MyData.MyParameter[]
            {
                new MyData.MyParameter("@DateBegin",NowTime.AddDays(-1).Date, MyData.MyParameter.MyDataType.DateTime),
                new MyData.MyParameter("@DateEnd",NowTime.Date , MyData.MyParameter.MyDataType.DateTime)
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
                        string mailto = ConfigurationManager.AppSettings["RejectMailTo"], mailcc = ConfigurationManager.AppSettings["RejectMailCC"];
                        MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        sm.MailTo = mailto;
                        //"jane123@my.imedia.com.tw,lwy@my.imedia.com.tw";
                        sm.MailCC = mailcc; //"jenny@imedia.com.tw,sparktsai@my.imedia.com.tw,my80@my.imedia.com.tw";
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

        #region 发送10天前没有结案的工单

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
<DIV><FONT size=3 face=PMingLiU>{3}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容为十天前（{0:yy/MM/dd}前）所有未结且交货期小于今日的工单列表。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; （说明：工单未结是指—工单需要入库且入库完成没有打勾的工单。）</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<TABLE style=""BORDER-COLLAPSE: collapse"" cellSpacing=0 cellPadding=0 width=""100%"" border=0>
  <TBODY>
    <TR>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    工单号    
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    订单号
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    客户编号
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    工单性质    
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    产品类别
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    产品编号
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    产品料号
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    工单数量
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    完成数量
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    入库数量
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    输入日期
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    输入人
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    审核日期
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    审核人
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    工单交期
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    刻交期人
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    交期说明
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    核销人
    </TD>
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    核销说明
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

                string SQL = @" 
Select RdsNo,OrderNo,CustID,Code,pDeliver as SendDate,pNumb,InputDate,Inputer,CheckDate,Checker,Remark,finishDate,Property as PropertyID,
       FinishRemark,FinishMan,SRemark as SendRemark,RMan as SendMan,RTime as SendDateSignTime,InStockFinishNumb,
	   Name=Convert(nVarchar(100),Null),Size=Convert(nVarchar(100),Null),
	   TypeName=Convert(nVarchar(100),null),PropertyName=(Select Name from [moProdProperty] Where Code=a.Property),
	   FinishNumb=(Select Top 1 FinishNumb from moProdProcedure Where a.ProdProcID = ProcNo And zbid=a.id)
Into #T
from moProduce a 
Where StockDate is Null And Year(InputDate)>2013 And CheckDate < DateAdd(dd,-9,GetDate()) And
      pDeliver < DateAdd(dd,1,GetDate()) And isNull(StopStock,0)=0
Update a Set a.Name = b.FullName,a.TypeName=b.mTypeName,a.Size=b.Size From #T a,AllMaterialView b Where a.Code=b.Code
Update a Set a.InStockFinishNumb=(Select Sum(Numb) from stPrdStocklst b Where a.RdsNo=b.ProductNo) From #T a
Select * from #T Order by InputDate Asc
Drop Table #T
";
                DateTime NowTime = DateTime.Now;
                string brs = "";
                MyRecord.Say("后台计算");
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL))
                {
                    MyRecord.Say("工单已经获取");
                    if (md != null && md.MyRows.Count > 0)
                    {
                        foreach (var ri in md.MyRows)
                        {
                            string sfNumbWord = Convert.ToString(ri["InStockFinishNumb"]), spNumbWord = Convert.ToString(ri["pNumb"]), sffNumbWord = Convert.ToString(ri["FinishNumb"]);
                            int sfNumb = 0; if (!int.TryParse(sfNumbWord, out sfNumb)) sfNumb = 0;
                            int spNumb = 0; if (!int.TryParse(spNumbWord, out spNumb)) spNumb = 0;
                            int sffNumb = 0; if (!int.TryParse(sffNumbWord, out sffNumb)) sffNumb = 0;

                            brs += string.Format(br, Convert.ToString(ri["RdsNo"]),
                                                    Convert.ToString(ri["OrderNo"]),
                                                    Convert.ToString(ri["CustID"]),
                                                    Convert.ToString(ri["PropertyName"]),
                                                    Convert.ToString(ri["TypeName"]),
                                                    Convert.ToString(ri["Code"]),
                                                    Convert.ToString(ri["Name"]),
                                                    spNumb,
                                                    sffNumb,
                                                    sfNumb,
                                                    string.Format("{0:yy/MM/dd HH:mm}", Convert.ToDateTime(ri["InputDate"])),
                                                    Convert.ToString(ri["Inputer"]),
                                                    string.Format("{0:yy/MM/dd HH:mm}", Convert.ToDateTime(ri["CheckDate"])),
                                                    Convert.ToString(ri["Checker"]),
                                                    string.Format("{0:yy/MM/dd HH:mm}", Convert.ToDateTime(ri["SendDate"])),
                                                    Convert.ToString(ri["SendMan"]),
                                                    Convert.ToString(ri["SendRemark"]),
                                                    Convert.ToString(ri["FinishMan"]),
                                                    Convert.ToString(ri["FinishRemark"])
                                                    );
                        }
                        MyRecord.Say(string.Format("表格一共：{0}行，表格已经生成。", md.Rows.Count));
                        MyRecord.Say("创建SendMail。");
                        MyBase.SendMail sm = new MyBase.SendMail();
                        MyRecord.Say("加载邮件内容。");
                        sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, NowTime.AddDays(-10), brs, NowTime, MyBase.CompanyTitle));
                        sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}_10日未结工单提醒。", NowTime, MyBase.CompanyTitle));
                        string mailto = ConfigurationManager.AppSettings["TenDayMailTo"], mailcc = ConfigurationManager.AppSettings["TenDayMailCC"];
                        MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        sm.MailTo = mailto; // "my18@my.imedia.com.tw,xang@my.imedia.com.tw,lghua@my.imedia.com.tw,my64@my.imedia.com.tw";
                        sm.MailCC = mailcc; // "jane123@my.imedia.com.tw,lwy@my.imedia.com.tw,my80@my.imedia.com.tw";
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
                RdsNo = Convert.ToString(r["RdsNo"]);
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
                Property = Convert.ToString(r["Property"]);
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
            public string Property { get; set; }
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
<DIV><FONT size=3 face=PMingLiU>{4}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 请注意，以下内容为{0:yy/MM/dd HH:mm}至{1:yy/MM/dd HH:mm}之间完工有差异的工单。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; (为了减小邮件大小，表格中只包含有差异数的工序。详细情况请在ERP生产进度中查询。)</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
{2}</FONT>
</DIV>
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{3:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
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
    <TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
    模數
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
    </TR>
";
                #endregion
                MyRecord.Say("1.设定数据条件，准备开始计算。");
                DateTime NowTime = DateTime.Now;
                DateTime beginTime = NowTime.AddMonths(-6).AddDays(1 - NowTime.Day).Date, endTime = NowTime.AddDays(-10).Date.AddMilliseconds(-10);
                DateTime stTime = new DateTime(2015, 11, 01).Date;
                if (beginTime < stTime)
                {
                    beginTime = stTime;
                }
                #region 加载数据源
                Thread.Sleep(200);
                MyData.MyParameter[] mps = new MyData.MyParameter[]
                    {
                        new MyData.MyParameter("@BTime",beginTime, MyData.MyParameter.MyDataType.DateTime),
                        new MyData.MyParameter("@ETime",endTime, MyData.MyParameter.MyDataType.DateTime),
                        new MyData.MyParameter("@ProduceRdsNo",string.Empty),
                        new MyData.MyParameter("@PName",string.Empty),
                        new MyData.MyParameter("@PCode",string.Empty),
                        new MyData.MyParameter("@CUST",string.Empty),
                        new MyData.MyParameter("@FinishStatus",0, MyData.MyParameter.MyDataType.Int),
                        new MyData.MyParameter("@StockStatus",3, MyData.MyParameter.MyDataType.Int),
                        new MyData.MyParameter("@ProdType",string.Empty),
                        new MyData.MyParameter("@PhoneSubject",string.Empty),
                        new MyData.MyParameter("@BK",-1, MyData.MyParameter.MyDataType.Int),
                        new MyData.MyParameter("@Bond",-1, MyData.MyParameter.MyDataType.Int)
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
                                orderby Convert.ToString(a["ProduceRdsNo"])
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
                    MyRecord.Say("2.在本机计算表格内容。");
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
                                                pickedNumber = vPi.ReqNumb;
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
                            fWord = MyConvert.ZHLC(string.Format("有差异 工序：{0}，机台：{1}，差异数：{2}", gi.procname, gi.machinename, gi.defnumb));
                        }
                        else
                        {
                            fWord = MyConvert.ZHLC("无差异");
                        }
                        MyRecord.Say(string.Format("正在处理：{0} 耗时：{1:0.0}秒，{2}", xNote.RdsNo, (DateTime.Now - bTime).TotalSeconds, fWord));
                    }

                    #endregion
                    #region 筛选加工
                    MyRecord.Say("对数据进行筛选加工。");
                    List<GridItem> _ThisGridDataSource = new List<GridItem>();
                    DateTime tmDateTime = new DateTime(1998, 1, 1);
                    _ThisGridDataSource = GridDataSource;
                    var nvk = from a in _ThisGridDataSource
                              where a.defnumb != 0 && a.FinishDate > tmDateTime && a.InStockNumb > 0 && !a.Finalized
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
                                                     item.partid,
                                                     item.DeptmantName,
                                                     item.procname,
                                                     item.machinename,
                                                     item.picknumb,
                                                     item.finishnumb,
                                                     item.rejectnumb,
                                                     item.wastagenumb,
                                                     item.defnumb,
                                                     item.finishword,
                                                     string.Format("{0:0.00%}", item.yield),
                                                     item.ColNumb
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
                    if (EmailGridRow > 0)
                    {
                        MyRecord.Say("创建SendMail。");
                        MyBase.SendMail sm = new MyBase.SendMail();
                        MyRecord.Say("加载邮件内容。");
                        bodyGrid = EmailGridRow > 0 ? string.Format(bodyGrid, brs) : "没有找到差异内容。";
                        sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, beginTime, endTime, bodyGrid, NowTime, MyBase.CompanyTitle));
                        sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yy年MM月dd日}結單差異數提醒。", NowTime, MyBase.CompanyTitle));
                        string mailto = ConfigurationManager.AppSettings["FinishMailTo"], mailcc = ConfigurationManager.AppSettings["FinishMailCC"];
                        MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        sm.MailTo = mailto;  //"my18@my.imedia.com.tw,xang@my.imedia.com.tw,my64@my.imedia.com.tw,cui@my.imedia.com.tw,my33@my.imedia.com.tw,my81@my.imedia.com.tw,my29@my.imedia.com.tw,yang@my.imedia.com.tw,my62@my.imedia.com.tw,my67@my.imedia.com.tw,my98@my.imedia.com.tw";
                        sm.MailCC = mailcc;  //"lwy@my.imedia.com.tw,jane123@my.imedia.com.tw,jenny@imedia.com.tw,sparktsai@my.imedia.com.tw,my80@my.imedia.com.tw";
                        //sm.MailTo = "my80@my.imedia.com.tw";
                        MyRecord.Say("发送邮件。");
                        NowTime = DateTime.Now;
                        sm.SendOut();
                        MyRecord.Say(string.Format("发送完成，耗時：{0:#,##0.00}", (DateTime.Now - NowTime).TotalSeconds));
                    }
                    else
                    {
                        MyRecord.Say("没有找到资料。");
                    }
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
                            Select Distinct ProduceNo as RdsNo from ProdDailyReport Where InputDate Between @InputBegin And @InputEnd 
                            Union
                            Select Distinct ProductNo as RdsNo from stOutProdlst Where InputDate Between @InputBegin And @InputEnd
                            Union
                            Select Distinct ProduceRdsNo as RdsNo from stOtherOutlst Where CreateDate Between @InputBegin And @InputEnd
                            ) T
                            Where RdsNo is Not Null
                            Order by RdsNo
                        ";
                MyData.MyDataTable mTableProduceFinished = new MyData.MyDataTable(SQL,
                    new MyData.MyParameter("@InputBegin", StartTime, MyData.MyParameter.MyDataType.DateTime),
                    new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime)
                    );
                if (_StopAll) return;
                if (mTableProduceFinished != null && mTableProduceFinished.MyRows.Count > 0)
                {
                    string memo = string.Format("读取了{0}条记录，下面开始计算....", mTableProduceFinished.MyRows.Count);
                    MyRecord.Say(memo);
                    int mTableProduceFinishedCount = 1;
                    foreach (MyData.MyDataRow r in mTableProduceFinished.MyRows)
                    {
                        if (_StopAll) return;
                        string RdsNo = Convert.ToString(r["RdsNo"]);
                        DateTime mStartTime = DateTime.Now;
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
                MyData.MyParameter mp = new MyData.MyParameter("@rdsno", CurrentRdsNO);
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

        #region 定时审核排程

        void ConfirmProcessPlan()
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
                string SQL = "";
                MyRecord.Say("---------------------启动定时审核排程。------------------------------");
                DateTime bStartTime = DateTime.Now;
                DateTime NowTime = DateTime.Now, TimeBegin = DateTime.MinValue, TimeEnd = DateTime.MinValue, TimeSet = DateTime.MinValue;
                MyRecord.Say("第一步，刷新所有部门和工序，最后一份排程。");
                if (NowTime.Hour == 10)
                {
                    TimeBegin = NowTime.Date.AddHours(7);
                    TimeEnd = NowTime.Date.AddHours(18);
                    TimeSet = NowTime.Date.AddHours(20);
                }
                else if (NowTime.Hour == 22)
                {
                    TimeBegin = NowTime.Date.AddHours(19);
                    TimeEnd = NowTime.Date.AddDays(1).Date.AddHours(6);
                    TimeSet = NowTime.Date.AddDays(1).Date.AddHours(8);
                }
                else
                {
                    return;
                }

                SQL = @"Select Max(_ID) as ID from [_PMC_ProdPlan] a Where 
a.PlanBegin Between @TimeBgein And @TimeEnd And PlanEnd is Null 
And Not Exists(Select Top 1 _ID From [_PMC_ProdPlan] b Where b.PlanBegin Between @TimeBgein And @TimeEnd And PlanEnd is Not Null And a.Process=b.Process And a.Department=b.Department)
Group by Process,Department ";
                MyData.MyDataTable mLastPlan = new MyData.MyDataTable(SQL,
                    new MyData.MyParameter("@TimeBgein", TimeBegin, MyData.MyParameter.MyDataType.DateTime),
                    new MyData.MyParameter("@TimeEnd", TimeEnd, MyData.MyParameter.MyDataType.DateTime));
                if (_StopAll) return;
                if (mLastPlan != null && mLastPlan.MyRows.Count > 0)
                {
                    string memo = string.Format("读取了{0}条记录，下面开始计算....", mLastPlan.MyRows.Count);
                    MyRecord.Say(memo);
                    int mLastPlanCount = 1;
                    foreach (MyData.MyDataRow r in mLastPlan.MyRows)
                    {
                        if (_StopAll) return;
                        MyRecord.Say(string.Format("审核第{0}条:", mLastPlanCount));
                        int PlanID = Convert.ToInt32(r["ID"]);
                        DateTime mStartTime = DateTime.Now;
                        ConfirmAndCheckPlan(PlanID, TimeSet);
                        MyRecord.Say(string.Format("耗时：{0:#,#0.00}秒。", (DateTime.Now - mStartTime).TotalSeconds));
                        mLastPlanCount++;
                    }
                }
                else
                {
                    MyRecord.Say("没有获取到任何内容");
                }
                MyRecord.Say(string.Format("计算完毕，一共耗时：{0:#,##0.00}分钟。", (DateTime.Now - bStartTime).TotalSeconds));
                MyRecord.Say("-----------------------审核完成。-------------------------------");
            }
            catch (Exception ex)
            {
                MyRecord.Say(ex);
            }
        }

        bool ConfirmAndCheckPlan(int xPlanID, DateTime PlanEndSet)
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
    Values(@zbid,'系統審核',+1,'',@rdsno,'審核',2,1,1)
End
Update [_PMC_ProdPlan] Set Status=2,Checker='系統審核',CheckDate=GetDate() Where [_ID]=@zbid
Insert Into [_PMC_ProdPlan_StatusRecorder]
        (zbid,Author,[state],memo,rdsno,type,typeid,CheckIn,CheckOut)
Values(@zbid,'系統審核',+1,'',@rdsno,'確認',2,1,1)
SET NOCOUNT OFF
";
            MyData.MyCommand mc = new MyData.MyCommand();
            return mc.Execute(SQL,
                new MyData.MyParameter("@zbid", xPlanID, MyData.MyParameter.MyDataType.Int),
                new MyData.MyParameter("@PlanEndSet", PlanEndSet, MyData.MyParameter.MyDataType.DateTime)
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
                    new MyData.MyParameter("@DateBegin", BDate, MyData.MyParameter.MyDataType.DateTime),
                    new MyData.MyParameter("@DateEnd", EDate, MyData.MyParameter.MyDataType.DateTime)
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
                        string mailto = ConfigurationManager.AppSettings["ProduceMailTo"], mailcc = ConfigurationManager.AppSettings["ProduceMailCC"];
                        MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        sm.MailTo = mailto; //"my18@my.imedia.com.tw,xang@my.imedia.com.tw,lghua@my.imedia.com.tw,my64@my.imedia.com.tw";
                        sm.MailCC = mailcc; //"jane123@my.imedia.com.tw,lwy@my.imedia.com.tw,my80@my.imedia.com.tw";
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
                        MyData.MyParameter mpid = new MyData.MyParameter("@ID", r["_ID"], MyData.MyParameter.MyDataType.Int);
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
                        MyData.MyParameter mpRdsNo = new MyData.MyParameter("@rdsno", RdsNo);
                        MyData.MyParameter mpid = new MyData.MyParameter("@ID", r["_ID"], MyData.MyParameter.MyDataType.Int);
                        MyData.MyParameter mpEndDate = new MyData.MyParameter("@PlanEnd", r["PlanEnd"], MyData.MyParameter.MyDataType.DateTime);
                        string SQLDetial = @"Delete from [_PMC_ProdPlan_List] Where zbid = @ID And Bdd > DateAdd(HH,3,@PlanEnd)";
                        string SQLMain = @"Update [_PMC_ProdPlan] Set Status=3 Where [_ID]=@ID";
                        string SQLStatus = @"Insert Into [_PMC_ProdPlan_StatusRecorder](zbid,Author,[state],memo,rdsno,type,typeid,CheckIn,CheckOut) Values(@ID,'自動審核',1,'關閉排程，到 ' + Convert(VarChar(100),DateAdd(HH,3,@PlanEnd),120),@rdsno,'關閉',2,1,1)";
                        MyData.MyParameter[] mpss = new MyData.MyParameter[]
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
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, new MyData.MyParameter("@DateEnd", EDate, MyData.MyParameter.MyDataType.DateTime)))
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
                        string mailto = ConfigurationManager.AppSettings["8DMailTo"], mailcc = ConfigurationManager.AppSettings["8DMailCC"];
                        MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        sm.MailTo = mailto;
                        sm.MailCC = mailcc;
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
                    new MyData.MyParameter("@InputBegin", StartTime, MyData.MyParameter.MyDataType.DateTime),
                    new MyData.MyParameter("@InputEnd", StopTime, MyData.MyParameter.MyDataType.DateTime)
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
                MyData.MyParameter mp = new MyData.MyParameter("@rdsno", CurrentRdsNO);
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
                MyRecord.Say("-----------------未结束维修申请-------------------------");
                MyRecord.Say("生成内容");

                string body = MyConvert.ZH_TW(@"
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
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, new MyData.MyParameter("@DateEnd", EDate, MyData.MyParameter.MyDataType.DateTime)))
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
                        string mailto = ConfigurationManager.AppSettings["MachineMailTo"], mailcc = ConfigurationManager.AppSettings["MachineMailCC"];
                        MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", mailto, mailcc));
                        sm.MailTo = mailto;
                        sm.MailCC = mailcc;
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
Update b Set b.InStockNumb=ISNULL((Select Sum(Numb) from coPurchStocklst c Where PurchNo=a.rdsno And PID = b.id),0) From coPurchase a,coPurchlst b Where a.id=b.zbid And b.InStockNumb is Null And a.status > 0
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

Update a Set a.ReqID=(Select n.ReqPID From coPurchase m,coPurchLst n Where m.id=n.zbid And m.RdsNo=a.PurchNo And n.id = a.pid)
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

        #region 计算库存的平均周转天数和最后出入库时间，价格，反馈入库日期入库信息到出库

        void StockCalculateLoader()
        {
            MyRecord.Say("开启库存后台计算..........");
            Thread t = new Thread(new ThreadStart(StockCalculate));
            t.IsBackground = true;
            t.Start();
            MyRecord.Say("开启库存后台计算成功。");
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

Select a.Code,Min(b.InStockDate) StkDate Into #ST From [_ST_StockListByLotNoView] a Inner Join [_WH_InStockLotNo_View] b ON a.LotNo = b.LotNo Group by a.Code

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

Update pbMaterial Set LastOutStockDate = (Select StkDate From #T Where Code=pbMaterial.Code)

Update pbKzmsg Set LastOutStockDate = (Select StkDate From #T Where Code=pbKzmsg.Code)

Update pbProduct Set LastOutStockDate = (Select StkDate From #T Where Code=pbProduct.Code)


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

Update pbMaterial Set InventoryTurnOverDays = (Select StkDate From #X Where Code=pbMaterial.Code)

Update pbKzmsg Set InventoryTurnOverDays = (Select StkDate From #X Where Code=pbKzmsg.Code)

Update pbProduct Set InventoryTurnOverDays = (Select StkDate From #X Where Code=pbProduct.Code)

Drop Table #S
Drop Table #T
Drop Table #X

Set NoCount Off
";
                mcd.Add(SQL3, "SQL_3");
                ///修改出库的入库时间
                string SQL4 = @"
Set NoCount ON

Update stOutProdlst Set InStockDate = (Select InStockDate From [_WH_InStockLotNo_View] b Where b.LotNo=stOutProdlst.BatchNo) Where InStockDate is Null

Update stOtherOutlst Set InStockDate = (Select InStockDate From [_WH_InStockLotNo_View] b Where b.LotNo=stOtherOutlst.BatchNo) Where InStockDate is Null

Update stAdjustStocklst Set InStockDate = (Select InStockDate From [_WH_InStockLotNo_View] b Where b.LotNo=stAdjustStocklst.BatchNo) Where InStockDate is Null

Update coShiplst Set InStockDate = (Select InStockDate From [_WH_InStockLotNo_View] b Where b.LotNo=coShiplst.BatchNo) Where InStockDate is Null

Set NoCount Off
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
    }
}
