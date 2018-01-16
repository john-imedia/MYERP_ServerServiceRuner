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
using MYERP_ServerServiceRuner.Base;
using NPOI;
using NPOI.Util;

namespace MYERP_ServerServiceRuner
{
	public partial class MainService
	{
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
	   b.adjustTime,
	   b.HrNumb,
	   b.PartNO as PartID,
	   b.ProdNo,
	   b.ProdCode as ProductCode,
	   b.Remark,
	   b.PN as ColNumb,
	   b.Side,
	   b.AutoDuplexPrint,
	   MachineName = m.Name,
	   m.FinishMeasurement,m.StaticMeasurement,
	   ProcessName = mp.name,
	   ProductName = d.name,
	   DepartmentName = p.name,p.FullSortID,
	   ProdSide=(Case When b.ProcNo='2000' Then (Select Case When (ca1+ca2)>0 And (Cb1+Cb2)>0 Then 2 Else 1 End from moProdProcedure mp Where mp.zbid=b.PRODID And mp.ID=b.ProcID) Else 1 END)
from _PMC_ProdPlan a Inner Join _PMC_ProdPlan_List b ON a.[_ID]=b.zbid
					Left Outer Join moMachine m ON b.MachineCode = m.Code
					Left Outer Join pbProduct d On b.ProdCode = d.Code
					Left Outer Join moProcedure mp On a.Process = mp.Code
					Left Outer Join pbDept p On a.[Department] = p.[_id]
Where a.[Status] >=2 And isNull(b.YieldOff,0)=0
  And a.PlanBegin > @DateBegin And a.PlanEnd < @DateEnd And b.Edd< DateAdd(mi,@PlanExtendTime,a.PlanEnd) And b.Bdd > '2000-01-01 00:00:00'
Order by a.Department,a.Process,b.MachineCode,b.[_id]
";
			string SQLFinish = @"
Select ProcessID as ProcessCode,Numb1,MachinID as MachineCode,StartTime,EndTime,ProduceNo,Inputdate,Inputer,Numb2,Operator,Remark,
	   PartID,Remark2,ColNumb,NAColNumb,AdjustNumb,SampleNumb,Rejector,RejectDate,PaperColNumb,ProductCode,AccNoteRdsNo,Side,AutoDuplexPrint,DestroySubLevelItemCode
 from prodDailyReport where RptDate Between @fBeginDate And @fEndDate
";
			DateTime TimeBegin = DateTime.MinValue, TimeEnd = DateTime.MinValue, fTimeBegin = DateTime.MinValue, fTimeEnd = DateTime.MinValue;
			//if (classtype == 1) //白班
			//{
			//    fTimeBegin = NowTime.Date.AddHours(7).AddMinutes(55);
			//    fTimeEnd = NowTime.Date.AddDays(3).AddHours(22);
			//    TimeBegin = NowTime.Date.AddHours(7).AddMinutes(55);
			//    TimeEnd = NowTime.Date.AddHours(22).AddMinutes(15);
			//}
			//else  //夜班
			//{
			//    fTimeBegin = NowTime.Date.AddHours(19).AddMinutes(55);
			//    fTimeEnd = NowTime.Date.AddDays(4).AddHours(10);
			//    TimeBegin = NowTime.Date.AddHours(19).AddMinutes(55);
			//    TimeEnd = NowTime.Date.AddDays(1).AddHours(8).AddMinutes(15);
			//}

			if (classtype == 1) //白班
			{
				fTimeBegin = NowTime.Date.AddHours(6).AddMinutes(55);
				fTimeEnd = NowTime.Date.AddDays(3).AddHours(22);
				TimeBegin = NowTime.Date.AddHours(6).AddMinutes(55);
				TimeEnd = NowTime.Date.AddHours(22).AddMinutes(15);
			}
			else  //夜班
			{
				fTimeBegin = NowTime.Date.AddHours(18).AddMinutes(55);
				fTimeEnd = NowTime.Date.AddDays(4).AddHours(10);
				TimeBegin = NowTime.Date.AddHours(18).AddMinutes(55);
				TimeEnd = NowTime.Date.AddDays(1).AddHours(8).AddMinutes(15);
			}
			MyData.MyDataParameter[] mfps = new MyData.MyDataParameter[]
			{
				new MyData.MyDataParameter("@DateBegin", TimeBegin, MyData.MyDataParameter.MyDataType.DateTime),
				new MyData.MyDataParameter("@DateEnd",TimeEnd , MyData.MyDataParameter.MyDataType.DateTime),
				new MyData.MyDataParameter("@fBeginDate", fTimeBegin, MyData.MyDataParameter.MyDataType.DateTime),
				new MyData.MyDataParameter("@fEndDate", fTimeEnd, MyData.MyDataParameter.MyDataType.DateTime),
				new MyData.MyDataParameter("@PlanExtendTime",_PlanExtendTime , MyData.MyDataParameter.MyDataType.Int)
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
				#region 旧的关掉
				/*
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
						if (curItem.ProcessCode == "2000")
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
							if (curItem.FinishMeasurement == 1)
								PlanProdNumb = vd.Sum(m => m.ReqNumb);
							else
								PlanProdNumb = vd.Sum(m => m.ReqNumb * m.ColNumb);

							if (vf != null && vf.Count() > 0)
							{
								FinishSheetNumb = vf.Sum(x => x.FinishSheetNumb);
								if (curItem.FinishMeasurement == 1)
									FinishProdNumb = Math.Ceiling(vf.Sum(x => (x.FinishProdNumb / x.ColNumb)));
								else
									FinishProdNumb = vf.Sum(m => m.FinishProdNumb);
							}
						}
					}
				}
				 */
				#endregion
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
					int classType = curItem.classtype;
					hrNumb = g.Max(x => x.HrNumb);
					Inputer = curItem.Inputer;
					Checker = curItem.Checker;
					WorkLevel = g.Max(x => x.WorkLevel);
					var vd = from a in g
							 where (a.Bdd - a.PlanBegin).TotalMinutes > -5 && (a.Edd - a.PlanEnd).TotalMinutes < 5
							 select a;
					if (vd != null && vd.Count() > 0)
					{
						BDD = vd.Min(x => x.Bdd);
						EDD = vd.Max(x => x.Edd);
                        double LossTime = vd.Sum(x => x.AdjustTime + x.PrepTime), AllTime = vd.Sum(x => x.ReqTime);
                        PlanActivationTime = (AllTime - LossTime) / 60.00;
                        PlanAdjustmentTime = LossTime / 60.00;

						if (curItem.ProcessCode == "2000")
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

							var vr = from a in g
									 where a.FinishRejectSheetNumb > 0
									 select a;
							if (vr.IsNotEmptySet())
							{
								RejectProdNumb = vf.Sum(x => x.FinishRejectSheetNumb);
								RejectSheetNumb = vf.Sum(x => x.FinishRejectSheetNumb);
							}
						}
						else
						{
							var vf = from a in g
									 where a.Side == 0 && a.FinishSheetNumb > 0
									 select a;
							var vr = from a in g
									 where a.FinishRejectSheetNumb > 0
									 select a;

							PlanCount = vd.Count();
							FinishCount = vf.Count();

							PlanSheetNumb = vd.Sum(x => x.ReqNumb);
							if (curItem.FinishMeasurement == 1)
								PlanProdNumb = vd.Sum(m => m.ReqNumb);
							else
								PlanProdNumb = vd.Sum(m => m.ReqNumb * m.ColNumb);

							if (vf != null && vf.Count() > 0)
							{
								FinishSheetNumb = vf.Sum(x => x.FinishSheetNumb);
								if (curItem.FinishMeasurement == 1)
									FinishProdNumb = Math.Ceiling(vf.Sum(x => (x.FinishProdNumb / x.ColNumb)));
								else
									FinishProdNumb = vf.Sum(m => m.FinishProdNumb);
							}
							if (vr.IsNotEmptySet())
							{
								RejectSheetNumb = vf.Sum(x => x.FinishRejectSheetNumb);
								if (curItem.FinishMeasurement == 1)
									RejectProdNumb = Math.Ceiling(vf.Sum(x => (x.FinishRejectProdNumb / x.ColNumb)));
								else
									RejectProdNumb = vf.Sum(m => m.FinishRejectProdNumb);
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
			public double WorkLevel { get; set; }
			public double PlanCapbility { get; set; }
			public double RealCapbility { get; set; }
			public double RejectSheetNumb { get; set; }
			public double RejectProdNumb { get; set; }
			public double PlanActivationTime { get; set; }
            public double PlanAdjustmentTime { get; set; }
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
				AdjustTime = Convert.ToDouble(r["AdjustTime"]);
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
				FinishMeasurement = Convert.ToInt32(r["FinishMeasurement"]);
				StaticMeasurement = Convert.ToInt32(r["StaticMeasurement"]);
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
			public double AdjustTime { get; set; }
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
			public int StaticMeasurement { get; set; }
			public double FinishRejectSheetNumb { get; set; }
			public double FinishRejectProdNumb { get; set; }
			public int classtype { get; set; }
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
				DestroySubLevelItemCode = Convert.ToString(r["DestroySubLevelItemCode"]);
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

			public int LossNumb
			{
				get
				{
					return AdjustNumb + SampleNumb;
				}
			}

			public int LossProdNumb
			{
				get
				{
					return (AdjustNumb + SampleNumb) * PaperColNumb;
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

			public string DestroySubLevelItemCode { get; set; }



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
			MyRecord.Say("3.2 计算达成率和加载表格...\r\n                       设定各个条件。");
			DateTime fTimeBegin = DateTime.MinValue, fTimeEnd = DateTime.MinValue;
			if (classtype == 1) //白班
			{
				fTimeBegin = NowTime.Date.AddHours(6).AddMinutes(55);
				fTimeEnd = NowTime.Date.AddDays(3).AddHours(22);
			}
			else  //夜班
			{
				fTimeBegin = NowTime.Date.AddHours(18).AddMinutes(55);
				fTimeEnd = NowTime.Date.AddDays(4).AddHours(10);
			}
			List<Plan_GridItem> _GridData = new List<Plan_GridItem>();
			MyRecord.Say(string.Format("3.2  完工单限定时间：fTimeBegin={0}，fTimeEnd={1}", fTimeBegin, fTimeEnd));
			#region 印刷达成率
			MyRecord.Say("3.3 计算印刷排程达成率。");
			var vPrintProduceSource = from a in Plan_Lists
									  where (a.Bdd - a.PlanBegin).TotalMinutes > -5 && a.ProcessCode == "2000"
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

				var vpme = from a in _LocalPringSource
						   where a.MachineCode == item.MachineCode && a.DepartmentID == item.DepartmentID && a.ProcessCode == item.ProcessCode
						   select a.Edd;
				DateTime EndTime = vpme.Max();

				if (item.PlanBegin.Date < DateTime.Parse("2017-04-12"))
					EndTime = item.PlanEnd;
				else
					EndTime = EndTime > item.PlanEnd ? item.PlanEnd : EndTime;

				if (item.AutoDuplexPrint && item.ProdSide == 2 && item.Side == 0)
				{
					var vAutoPrint = from a in _LocalPringSource
									 where a.MachineCode == item.MachineCode && a.DepartmentID == item.DepartmentID && a.ProcessCode == item.ProcessCode &&
										   a.ProduceRdsNo == item.ProduceRdsNo && a.AutoDuplexPrint && a.ProdSide == 2 && a.Side == 1 && a.Edd > item.Edd
									 select a.Edd;
					if (vAutoPrint.IsNotEmptySet()) item.Edd = vAutoPrint.Min();
				}

				var vf = from a in Plan_FinishLists
						 where a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID && a.Side == item.Side && a.ProcessCode == "2000" && a.MachineCode == item.MachineCode
                            && a.InputDate >= fTimeBegin && a.InputDate <= fTimeEnd
							&& a.StartTime >= StartTime.AddMinutes(-1) && a.EndTime <= EndTime.AddMinutes(5)
                            && a.StartTime <= a.InputDate //完工单不能提前输入，开始时间必须小于输入时间 
						 select a;
				double finishProdNumb = vf.Sum(x => x.FinishProdNumb), finishNumb = vf.Sum(x => x.FinishNumb);
				if (item.StaticMeasurement == 1)
				{
					finishProdNumb = vf.Sum(x => (x.FinishProdNumb + x.RejectProdNumb + x.LossProdNumb));
					finishNumb = vf.Sum(x => x.FinishNumb + x.RejectNumb + x.LossNumb);
				}

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
				//if (finishNumb > (item.ReqNumb * 1.10) && item.ReqNumb >= 500)   //印刷超数量，要按照需求数计算，不可以超出110%
				//{
				//    item.FinishProdNumb = item.ReqNumb * item.ColNumb * 1.10;
				//    item.FinishSheetNumb = item.ReqNumb * 1.10;
				//}
				//else
				//{
				item.FinishProdNumb = finishProdNumb;
				item.FinishSheetNumb = finishNumb;
				//}
				item.Type = ((item.Bdd - item.PlanBegin).TotalMinutes > -5 && (item.Edd - item.PlanEnd).TotalMinutes < 5) ? MyConvert.ZHLC("正常") : MyConvert.ZHLC("超出");
				if (finishProdNumb > 0) foreach (var iv in vf) iv.iCalItem = true;
			}
			MyRecord.Say("3.3.2 完工数已经计算，计算印刷达成率");
			var vPrintGrid = from a in _LocalPringSource
							 where (!(a.AutoDuplexPrint && a.ProdSide == 2 && a.Side == 1))
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

				var vpme = from a in Plan_Lists
						   where a.MachineCode == item.MachineCode && a.DepartmentID == item.DepartmentID && a.ProcessCode == item.ProcessCode
						   select a.Edd;
				DateTime EndTime = vpme.Max();
				if (item.PlanBegin.Date < DateTime.Parse("2017-04-12"))
					EndTime = item.PlanEnd;
				else
					EndTime = EndTime > item.PlanEnd ? item.PlanEnd : EndTime;


				var vf = from a in Plan_FinishLists
						 where a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID && a.ProcessCode == item.ProcessCode && a.MachineCode == item.MachineCode
                               && a.InputDate >= fTimeBegin && a.InputDate <= fTimeEnd
							   && a.StartTime >= StartTime.AddMinutes(-1) && a.EndTime <= EndTime.AddMinutes(15)
                               && a.StartTime <= a.InputDate //完工单不能提前输入，开始时间必须小于输入时间
						 select a;
				double finishProdNumb = vf.Sum(x => x.FinishProdNumb), finishNumb = vf.Sum(x => x.FinishNumb);

				if (item.StaticMeasurement == 1)
				{
					finishProdNumb = vf.Sum(x => (x.FinishProdNumb + x.RejectProdNumb + x.LossProdNumb));
					finishNumb = vf.Sum(x => x.FinishNumb + x.RejectNumb + x.LossNumb);
				}

				if (item.ProcessCode != "9035")
				{
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
				}
				else
				{
					var v9035 = from a in vf
								group a by a.DestroySubLevelItemCode into g
								select new
								{
									FinishProdNumb = g.Sum(x => (x.FinishProdNumb + x.RejectProdNumb)),
									FinishNumb = g.Sum(x => (x.FinishNumb + x.RejectNumb))
								};
					if (v9035 != null && v9035.Count() > 0)
					{
						finishProdNumb = v9035.Max(x => x.FinishProdNumb);
						finishNumb = v9035.Max(x => x.FinishNumb);
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
				MyRecord.Say("-----------开始计算排程达成率----------");
				MyRecord.Say("1.加载邮件体。");
				#region 表头
				string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{4}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {0:yy/MM/dd} {1} 所有排程及达成率。</FONT></DIV>
<DIV><FONT size=2 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; （说明：完工数只计算良品数；只计算计划开始时间至{3:MM/dd HH时}之间的完工单。）</FONT></DIV>
{2}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{3:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");
				string gd = @"
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
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	达成率
	</TD>
	</TR>
	{0}
</TBODY></TABLE></FONT>
</DIV>
";
				string sd = @"
<DIV><FONT size=5 face=PMingLiU color=#ff0080>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;没有安排生产排程。</FONT>
</DIV>
";
				#endregion
				MyRecord.Say("2.加载表格行。");
				#region 表格行
				string br = @"
	<TR>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{0}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{1}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{2}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{3}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{4}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{5}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{6}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{7}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{8}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{9}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{10}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{11}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: {13}"" align=center >
	{12}
	</TD>
	</TR>
";
				#endregion
				int classtype = 1;
				string byb = "";
				if (NowTime.Hour < 12 && NowTime.Hour > 5) //10点发送昨日夜班排程
				{
					classtype = 2;
					NowTime = NowTime.AddDays(-1);
					byb = "夜班";
				}
				else if (NowTime.Hour > 17 && NowTime.Hour < 23) //22点发送当日白天排程
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
					string xbd = string.Empty;
					if (vGridDataSource.Count() > 0)
					{
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
								string.Format("{0:0.00%}", y1),
								string.Format("{0:0.00%}", y1 * y2),
								y1 * y2 > 0.9 ? "transparent" : y1 * y2 > 0.75 ? "rgb(226, 234, 37)" : "rgb(243, 128, 153)"
								);
							MyRecord.Say(string.Format("4.1 - 第{0}行，部门：{1}，工序：{2}，机台：{3}，笔数达成率：{4}，产量达成率：{5}。", iRow, item.DepartmentName, item.ProcessName, item.MachineName, y1, y2));
							iRow++;
							xLine += xbr;
						}
						xbd = string.Format(gd, xLine);
					}
					else
					{
						xbd = sd;
					}
					MyRecord.Say(string.Format("表格一共：{0}行，已经生成。总耗时：{1}秒。", iRow, (DateTime.Now - t1).TotalSeconds));
					MyRecord.Say("创建SendMail。");
					MyBase.SendMail sm = new MyBase.SendMail();
					MyRecord.Say("加载邮件内容。");
					sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, NowTime, byb, xbd, DateTime.Now, MyBase.CompanyTitle));
					sm.Subject = MyConvert.ZH_TW(string.Format("{2}{0:yy年MM月dd日}{1}排程及达成率", NowTime, byb, MyBase.CompanyTitle));
					string MailTo = ConfigurationManager.AppSettings["PlanMailTo"], MailCC = ConfigurationManager.AppSettings["PlanMailCC"];
					MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", MailTo, MailCC));
					sm.MailTo = MailTo;
					sm.MailCC = MailCC;
					//sm.MailTo = "my80@my.imedia.com.tw";
					MyRecord.Say("发送邮件。");
					sm.SendOut();
					MyRecord.Say("已经发送。");
				}
			}
			catch (Exception ex)
			{
				MyRecord.Say(ex);
			}
			MyRecord.Say("-------计算排程达成率完成------");
		}

		void ProdKanbanSaveOnceLoader(DateTime NowTime)
		{
			_StopProdPlanForSaved = false;
			Thread tttt = new Thread(new ThreadStart(
				() =>
				{
					ProdPlanForSaveRunner(NowTime);
					KanbanRecorder();
				}));
			tttt.IsBackground = true;
			tttt.Start();
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
			MyRecord.Say(string.Format("向前推：{0}日。", CacluatePlanRateDaySpanTimes));
			DateTime xDate = ((DateTime)NowTime).AddDays(-CacluatePlanRateDaySpanTimes).Date;
			MyRecord.Say(string.Format("当前时间：{0}，计算开始时间：{1}", NowTime, xDate));
			for (int i = 0; i <= CacluatePlanRateDaySpanTimes; i++)
			{
				if (_StopProdPlanForSaved) return;
				DateTime iDate = xDate.AddDays(i).Date;
				MyRecord.Say("--------------------------------------------------------");
				MyRecord.Say(string.Format("计算：{0}，开始。", iDate));
				MyRecord.Say(string.Format("计算时间：{0}，白班", iDate));
				ProdPlanForSave(iDate, 1);
				Thread.Sleep(1000);
				if (_StopProdPlanForSaved) return;
				MyRecord.Say(string.Format("计算时间：{0}，夜班", iDate));
				ProdPlanForSave(iDate, 2);
				MyRecord.Say(string.Format("计算：{0}，完成。", iDate));
				MyRecord.Say("--------------------------------------------------------");
				Thread.Sleep(5000);
			}
		}

		void ProdPlanForSave(DateTime NowTime, int classtype)
		{
			try
			{
				string byb = "";
				MyData.MyCommand xmcd = new MyData.MyCommand();
				if (classtype == 1)
					byb = "白班";
				else if (classtype == 2)
					byb = "夜班";

				MyRecord.Say(string.Format("1.处理条件，NowTime={0}，班次：{1}，班次ID：{2}，开始计算数据源。", NowTime, byb, classtype));
				if (NowTime > DateTime.MinValue && classtype > 0)
				{
					MyRecord.Say(string.Format("2.删除当日当班达成率内容。PlanBegin = {0}，PlanType={1}", NowTime.Date, classtype));
					string SQLDeleteToday = "Delete From [_PMC_ProdPlan_YieldRate] Where Convert(VarChar(10),PlanBegin,121) = Convert(VarChar(10),@PlanBegin,121) And PlanType = @PlanType";
					MyData.MyDataParameter[] mpDelete = new MyData.MyDataParameter[]
					{
						new MyData.MyDataParameter("@PlanBegin",NowTime.Date, MyData.MyDataParameter.MyDataType.DateTime),
						new MyData.MyDataParameter("@PlanType",classtype, MyData.MyDataParameter.MyDataType.Int)
					};
					xmcd.Add(SQLDeleteToday, "DeleteAllDay", mpDelete);

					MyRecord.Say("3.获取并计算达成率");
					List<Plan_GridItem> _GridData = Plan_LoadFinishedRate(NowTime, classtype);
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
Insert Into [_PMC_ProdPlan_YieldRate](PlanRdsNo,PlanType,PlanBegin,PlanEnd,Inputer,Checker,DeaprtmentID,DepartmentHrNumb,ProcessCode,MachineCode,PlanSheetNumb,FinishSheetNumb,PlanNumb,FinishNumb,PlanProdNumb,FinishProdNumb,WorkLevel,PlanCapbility,RealCapbility)
							  Values(@PlanRdsNo,@PlanType,@PlanBegin,@PlanEnd,@Inputer,@Checker,@DeaprtmentID,@DepartmentHrNumb,@ProcessCode,@MachineCode,@PlanSheetNumb,@FinishSheetNumb,@PlanNumb,@FinishNumb,@PlanProdNumb,@FinishProdNumb,@WorkLevel,@PlanCapbility,@RealCapbility)";
						MyData.MyDataParameter[] amps = new MyData.MyDataParameter[]
						{
							new MyData.MyDataParameter("@PlanRdsNo",item.RdsNo),
							new MyData.MyDataParameter("@PlanType",classtype, MyData.MyDataParameter.MyDataType.Int ),
							new MyData.MyDataParameter("@PlanBegin",item.BDD, MyData.MyDataParameter.MyDataType.DateTime),
							new MyData.MyDataParameter("@PlanEnd",item.EDD, MyData.MyDataParameter.MyDataType.DateTime),
							new MyData.MyDataParameter("@Inputer",item.Inputer),
							new MyData.MyDataParameter("@Checker",item.Checker),
							new MyData.MyDataParameter("@DeaprtmentID",item.DepartmentID, MyData.MyDataParameter.MyDataType.Int),
							new MyData.MyDataParameter("@DepartmentHrNumb",item.hrNumb, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@ProcessCode",item.ProcessCode),
							new MyData.MyDataParameter("@MachineCode",item.MachineCode),
							new MyData.MyDataParameter("@PlanSheetNumb",item.PlanCount , MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@FinishSheetNumb",item.FinishCount , MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@PlanNumb",item.PlanSheetNumb , MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@FinishNumb",item.FinishSheetNumb , MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@PlanProdNumb",item.PlanProdNumb, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@FinishProdNumb",item.FinishProdNumb, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@WorkLevel",item.WorkLevel, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@PlanCapbility",item.PlanCapbility, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@RealCapbility",item.RealCapbility, MyData.MyDataParameter.MyDataType.Numeric)
						};
						xmcd.Add(SQLSave, string.Format("X_Insert{0}", iRow), amps);
						MyRecord.Say(string.Format("4.3-第{0}行，部门：{1}，工序：{2}，机台：{3}，笔数达成率：{4:0.00%}，产量达成率：{5:0.00%}", iRow, item.DepartmentName, item.ProcessName, item.MachineName, item.FinishCount / item.PlanCount, item.FinishSheetNumb / item.PlanSheetNumb));
						iRow++;
					}

					MyRecord.Say(string.Format("一共：{0}行，已经生成。总耗时：{1}秒。语句：{2}行。", iRow, (DateTime.Now - t1).TotalSeconds, xmcd.SQLCmdColl.Count));
					MyRecord.Say("开始提交，等2秒。");

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
			//MyRecord.Say("-----------保存排程达成率完成-----------");
		}

		#endregion

		#region 计算OEE达成率
		private List<Plan_GridItem> OEE_LoadFinishedRate(DateTime NowTime, int classtype)
		{
			MyRecord.Say("3.1 读取资料数据源,并保存到本机...");
			MyRecord.Say("3.1.1 读取数据源");
			Plan_LoadDataSource(NowTime, classtype);
			MyRecord.Say("3.2 计算达成率和加载表格...\r\n                       设定各个条件。");
			DateTime fTimeBegin = DateTime.MinValue, fTimeEnd = DateTime.MinValue;

			if (classtype == 1) //白班
			{
				fTimeBegin = NowTime.Date.AddHours(7).AddMinutes(55);
				fTimeEnd = NowTime.Date.AddDays(3).AddHours(22);
			}
			else  //夜班
			{
				fTimeBegin = NowTime.Date.AddHours(19).AddMinutes(55);
				fTimeEnd = NowTime.Date.AddDays(4).AddHours(10);
			}
			List<Plan_GridItem> _GridData = new List<Plan_GridItem>();
			MyRecord.Say(string.Format("3.2.1  完工单限定时间：fTimeBegin={0}，fTimeEnd={1}", fTimeBegin, fTimeEnd));
			#region 印刷达成率
			MyRecord.Say("3.3 计算印刷排程达成率。");
			var vPrintProduceSource = from a in Plan_Lists
									  where (a.Bdd - a.PlanBegin).TotalMinutes > -5 && a.ProcessCode == "2000"
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

				var vpme = from a in _LocalPringSource
						   where a.MachineCode == item.MachineCode && a.DepartmentID == item.DepartmentID && a.ProcessCode == item.ProcessCode
						   select a.Edd;
				DateTime EndTime = vpme.Max();

				if (item.PlanBegin.Date < DateTime.Parse("2017-04-12"))
					EndTime = item.PlanEnd;
				else
					EndTime = EndTime > item.PlanEnd ? item.PlanEnd : EndTime;

				if (item.AutoDuplexPrint && item.ProdSide == 2 && item.Side == 0)
				{
					var vAutoPrint = from a in _LocalPringSource
									 where a.MachineCode == item.MachineCode && a.DepartmentID == item.DepartmentID && a.ProcessCode == item.ProcessCode &&
										   a.ProduceRdsNo == item.ProduceRdsNo && a.AutoDuplexPrint && a.ProdSide == 2 && a.Side == 1 && a.Edd > item.Edd
									 select a.Edd;
					if (vAutoPrint.IsNotEmptySet()) item.Edd = vAutoPrint.Min();
				}

				var vf = from a in Plan_FinishLists
						 where a.InputDate >= fTimeBegin && a.InputDate <= fTimeEnd && a.ProcessCode == "2000" && a.MachineCode == item.MachineCode
							 && a.StartTime >= StartTime.AddMinutes(-1) && a.EndTime <= EndTime.AddMinutes(5)
							 && a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID && a.Side == item.Side
						 select a;
				double finishProdNumb = vf.Sum(x => x.FinishProdNumb), finishNumb = vf.Sum(x => x.FinishNumb),
					   rejectNumb = vf.Sum(x => x.RejectProdNumb / x.PaperColNumb + x.LossProdNumb / x.PaperColNumb), 
                       rejectProdNumb = vf.Sum(x => x.RejectProdNumb + x.LossProdNumb);
				finishProdNumb = vf.Sum(x => (x.FinishProdNumb + x.RejectProdNumb + x.LossProdNumb));
				finishNumb = vf.Sum(x => x.FinishProdNumb / x.PaperColNumb + x.RejectProdNumb / x.PaperColNumb + x.LossProdNumb / x.PaperColNumb);

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
						rejectNumb = 0;
						rejectProdNumb = 0;
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
				if (finishNumb > (item.ReqNumb * 1.05) && item.ReqNumb <= 500)   //印刷超数量，要按照需求数计算，不可以超出110%
				{
					item.FinishProdNumb = item.ReqNumb * item.ColNumb * 1.05;
					item.FinishSheetNumb = item.ReqNumb * 1.05;
				}
				else
				{
					item.FinishProdNumb = finishProdNumb;
					item.FinishSheetNumb = finishNumb;
				}
				item.FinishRejectSheetNumb = rejectNumb;
				item.FinishRejectProdNumb = rejectProdNumb;
				item.classtype = classtype;
				if (finishProdNumb > 0) foreach (var iv in vf) iv.iCalItem = true;
			}
			MyRecord.Say("3.3.2 完工数已经计算，计算印刷达成率");
			var vPrintGrid = from a in _LocalPringSource
							 where (!(a.AutoDuplexPrint && a.ProdSide == 2 && a.Side == 1))
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

				var vpme = from a in Plan_Lists
						   where a.MachineCode == item.MachineCode && a.DepartmentID == item.DepartmentID && a.ProcessCode == item.ProcessCode
						   select a.Edd;
				DateTime EndTime = vpme.Max();
				if (item.PlanBegin.Date < DateTime.Parse("2017-04-12"))
					EndTime = item.PlanEnd;
				else
					EndTime = EndTime > item.PlanEnd ? item.PlanEnd : EndTime;


				var vf = from a in Plan_FinishLists
						 where a.InputDate >= fTimeBegin && a.InputDate <= fTimeEnd && a.ProcessCode == item.ProcessCode && a.MachineCode == item.MachineCode
							   && a.StartTime >= StartTime.AddMinutes(-1) && a.EndTime <= EndTime.AddMinutes(15)
							   && a.ProduceRdsNo == item.ProduceRdsNo && a.PartID == item.PartID
						 select a;
				double finishProdNumb = vf.Sum(x => x.FinishProdNumb), finishNumb = vf.Sum(x => x.FinishNumb),
					   rejectNumb = vf.Sum(x => x.RejectProdNumb / x.PaperColNumb + x.LossProdNumb / x.PaperColNumb), rejectProdNumb = vf.Sum(x => x.RejectProdNumb + x.LossProdNumb);
				finishProdNumb = vf.Sum(x => (x.FinishProdNumb + x.RejectProdNumb + x.LossProdNumb));
				finishNumb = vf.Sum(x => x.FinishProdNumb / x.PaperColNumb + x.RejectProdNumb / x.PaperColNumb + x.LossProdNumb / x.PaperColNumb);

				if (item.ProcessCode != "9035")
				{
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
							rejectNumb = 0;
							rejectProdNumb = 0;
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
				}
				else
				{
					var v9035 = from a in vf
								group a by a.DestroySubLevelItemCode into g
								select new
								{
									FinishProdNumb = g.Sum(x => (x.FinishProdNumb + x.RejectProdNumb)),
									FinishNumb = g.Sum(x => (x.FinishNumb + x.RejectNumb))
								};
					if (v9035 != null && v9035.Count() > 0)
					{
						finishProdNumb = v9035.Max(x => x.FinishProdNumb);
						finishNumb = v9035.Max(x => x.FinishNumb);
					}
				}
				if (finishProdNumb <= 0) finishProdNumb = 0;
				if (finishNumb <= 0) finishNumb = 0;
				item.FinishProdNumb = finishProdNumb;
				item.FinishSheetNumb = finishNumb;
				item.FinishRejectSheetNumb = rejectNumb;
				item.FinishRejectProdNumb = rejectProdNumb;
				item.classtype = classtype;
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

		#region 保存OEE
		Thread t_OEE_ForSaveing;

		bool _Stop_OEE_ForSaved = false;

		void OEE_ForSaveLoader(DateTime NowTime)
		{
			if (t_OEE_ForSaveing != null && t_OEE_ForSaveing.IsAlive)
			{
				_Stop_OEE_ForSaved = true;
				DateTime xTime = DateTime.Now;
				do
				{
					_Stop_OEE_ForSaved = true;
				} while (t_OEE_ForSaveing.IsAlive && (DateTime.Now - xTime).TotalSeconds > 10);
			}
			_Stop_OEE_ForSaved = false;
			t_OEE_ForSaveing = new Thread(new ParameterizedThreadStart(OEE_ForSaveRunner));
			t_OEE_ForSaveing.IsBackground = true;
			t_OEE_ForSaveing.Start(NowTime);
		}

		void OEE_ForSaveRunner(object NowTime)
		{
			MyRecord.Say("--------------OEE计算------------------------------------------");
			MyRecord.Say(string.Format("向前推：{0}日。", CacluatePlanRateDaySpanTimes));
			DateTime xDate = ((DateTime)NowTime).AddDays(-CacluatePlanRateDaySpanTimes).Date;
			MyRecord.Say(string.Format("当前时间：{0}，计算开始时间：{1}", NowTime, xDate));
			for (int i = 0; i <= CacluatePlanRateDaySpanTimes; i++)
			{
				if (_Stop_OEE_ForSaved) return;
				DateTime iDate = xDate.AddDays(i).Date;
				MyRecord.Say("--------------------------------------------------------");
				MyRecord.Say(string.Format("计算：{0:yy/MM/dd} OEE，开始。", iDate));
				MyRecord.Say(string.Format("计算时间：{0}，白班", iDate));
				OEE_ForSave(iDate, 1);
				Thread.Sleep(1000);
				if (_Stop_OEE_ForSaved) return;
				MyRecord.Say(string.Format("计算时间：{0}，夜班", iDate));
				OEE_ForSave(iDate, 2);
				MyRecord.Say(string.Format("计算：{0}，完成。", iDate));
				MyRecord.Say("--------------------------------------------------------");
				Thread.Sleep(5000);
			}
			MyRecord.Say("--------------OEE计算全部完毕------------------------------------------");
		}

		void OEE_ForSave(DateTime NowTime, int classtype)
		{
			try
			{
				string byb = "";
				MyData.MyCommand xmcd = new MyData.MyCommand();
				if (classtype == 1) byb = "白班";
				else if (classtype == 2) byb = "夜班";
				MyRecord.Say(string.Format("1.处理条件，NowTime={0}，班次：{1}，班次ID：{2}，开始计算数据源。", NowTime, byb, classtype));
				if (NowTime > DateTime.MinValue && classtype > 0)
				{
					MyRecord.Say(string.Format("2.删除当日当班OEE内容。PlanBegin = {0}，PlanType={1}", NowTime.Date, classtype));
					string SQLDeleteToday = "Delete From [_PMC_ProdPlan_OEE] Where Convert(VarChar(10),PlanBegin,121) = Convert(VarChar(10),@PlanBegin,121) And PlanType = @PlanType";
					MyData.MyDataParameter[] mpDelete = new MyData.MyDataParameter[]
					{
						new MyData.MyDataParameter("@PlanBegin",NowTime.Date, MyData.MyDataParameter.MyDataType.DateTime),
						new MyData.MyDataParameter("@PlanType",classtype, MyData.MyDataParameter.MyDataType.Int)
					};
					xmcd.Add(SQLDeleteToday, "DeleteAllDay", mpDelete);

					MyRecord.Say("3.获取并计算OEE");
					List<Plan_GridItem> _GridData = OEE_LoadFinishedRate(NowTime, classtype);
					var vGridDataSource = from a in _GridData
										  where a.BDD > DateTime.Parse("2000-01-01") && a.EDD > DateTime.Parse("2000-01-01") && a.PlanCount > 0
										  orderby a.DepartmentFullSortID, a.ProcessCode, a.MachineCode, a.RdsNo
										  select a;
					MyRecord.Say("4.达成率计算完毕，生成保存语句。");
					string xLine = string.Empty; int iRow = 0; DateTime t1 = DateTime.Now;
					foreach (var item in vGridDataSource)
					{
						if (_Stop_OEE_ForSaved) return;
						double y1 = item.PlanCount != 0 ? Convert.ToDouble(item.FinishCount) / Convert.ToDouble(item.PlanCount) : 0;
						double y2 = item.FinishProdNumb != 0 ? Convert.ToDouble(item.FinishProdNumb) / Convert.ToDouble(item.PlanProdNumb) : 0;
						string SQLSave = @"
Delete from [_PMC_ProdPlan_OEE] Where PlanRdsNo=@PlanRdsNo And MachineCode=@MachineCode
Insert Into [_PMC_ProdPlan_OEE](PlanRdsNo,PlanType,PlanBegin,PlanEnd,DepartmentID,DepartmetHrNumb,ProcessCode,MachineCode,PlanNumb,FinishNumb,PlanProdNumb,FinishProdNumb,RejectNumb,RejectProdNumb,WorkLevel,PlanActivationTime,PlanAdjustmentTime)
							  Values(@PlanRdsNo,@PlanType,@PlanBegin,@PlanEnd,@DepartmentID,@DepartmetHrNumb,@ProcessCode,@MachineCode,@PlanNumb,@FinishNumb,@PlanProdNumb,@FinishProdNumb,@RejectNumb,@RejectProdNumb,@WorkLevel,@PlanActivationTime,@PlanAdjustmentTime)";
						MyData.MyDataParameter[] amps = new MyData.MyDataParameter[]
						{
							new MyData.MyDataParameter("@PlanRdsNo",item.RdsNo),
							new MyData.MyDataParameter("@PlanType",classtype, MyData.MyDataParameter.MyDataType.Int ),
							new MyData.MyDataParameter("@PlanBegin",item.BDD, MyData.MyDataParameter.MyDataType.DateTime),
							new MyData.MyDataParameter("@PlanEnd",item.EDD, MyData.MyDataParameter.MyDataType.DateTime),
							new MyData.MyDataParameter("@DepartmentID",item.DepartmentID, MyData.MyDataParameter.MyDataType.Int),
							new MyData.MyDataParameter("@DepartmetHrNumb",item.hrNumb, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@ProcessCode",item.ProcessCode),
							new MyData.MyDataParameter("@MachineCode",item.MachineCode),
							new MyData.MyDataParameter("@PlanNumb",item.PlanSheetNumb , MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@FinishNumb",item.FinishSheetNumb , MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@PlanProdNumb",item.PlanProdNumb, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@FinishProdNumb",item.FinishProdNumb, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@RejectNumb",item.RejectSheetNumb, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@RejectProdNumb",item.RejectProdNumb, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@WorkLevel",item.WorkLevel, MyData.MyDataParameter.MyDataType.Numeric),
							new MyData.MyDataParameter("@PlanActivationTime",item.PlanActivationTime, MyData.MyDataParameter.MyDataType.Numeric),
                            new MyData.MyDataParameter("@PlanAdjustmentTime",item.PlanAdjustmentTime, MyData.MyDataParameter.MyDataType.Numeric)
						};
						xmcd.Add(SQLSave, string.Format("X_Insert_{0}", iRow), amps);
						MyRecord.Say(string.Format("4.3-第{0}行，部门：{1}，工序：{2}，机台：{3}，笔数达成率：{4:0.00%}，产量达成率：{5:0.00%}", iRow, item.DepartmentName, item.ProcessName, item.MachineName, item.FinishCount / item.PlanCount, item.FinishSheetNumb / item.PlanSheetNumb));
						iRow++;
					}

					MyRecord.Say(string.Format("一共：{0}行，已经生成。总耗时：{1}秒。语句：{2}行。", iRow, (DateTime.Now - t1).TotalSeconds, xmcd.SQLCmdColl.Count));
					MyRecord.Say("开始提交，等2秒。");

					if (xmcd.Execute())
						MyRecord.Say(string.Format("保存{0:MM/dd}{1}达成率成功", NowTime, byb));
					else
						MyRecord.Say("运行出错。");
				}
			}
			catch (Exception ex)
			{
				MyRecord.Say(ex);
			}
		}

		#endregion

		#region 看板计算类
		public class ProductionScheduingRate
		{
			public class DateRateItem
			{
				/// <summary>
				/// 最高达成率比率。
				/// </summary>
				public const double TopPerformanceYield = 1.00;
				public DateRateItem(MyData.MyDataRow r)
				{
					PlanType = r.IntValue("PlanType");
					PlanBegin = r.DateTimeValue("PlanBegin");
					PlanEnd = r.DateTimeValue("PlanEnd");
					PlanRdsNo = r.Value("PlanRdsNo");
					DepartmentID = r.IntValue("DeaprtmentID");
					//DepartmentItem dept = Departments[DepartmentID];
					//if (dept.IsNotNull()) DepartmentName = dept.name;
					DepartmentName = r.Value("DepartmentName");
					DepartmentHrNumb = r.IntValue("DepartmentHrNumb");
					ProcessCode = r.Value("ProcessCode");
					ProcessName = r.Value("ProcessName");
					//ProcessItem process = Processes[ProcessCode];
					//if (process.IsNotNull()) ProcessName = process.Name;
					MachineCode = r.Value("MachineCode");
					//MachineItem machine = Machines[MachineCode];
					MachineName = r.Value("MachineName");
					PlanSheetNumb = r.Value<double>("PlanSheetNumb");
					FinishSheetNumb = r.Value<double>("FinishSheetNumb");
					PlanNumb = r.Value<double>("PlanNumb");
					FinishNumb = r.Value<double>("FinishNumb");
					PlanProdNumb = r.Value<double>("PlanProdNumb");
					FinishProdNumb = r.Value<double>("FinishProdNumb");
					DailyFinishNumb = r.Value<double>("DailyFinishNumb");
					DailyRejectNumb = r.Value<double>("DailyRejectNumb");
					RejectTimeNumb = r.Value<double>("RejectTimeNumb");
					isNotInspect = r.Value<double>("InspectScore2") == -100;// r["InspectScore"].IsNull();
					InspectScore = isNotInspect ? 1 : r.Value<double>("InspectScore");
					WeekName = r.IntValue("WeekName");
				}

				public int PlanType { get; set; }

				public string PlanTypeWord
				{
					get
					{
						return PlanType == 1 ? MyConvert.ZHLC("白班") :
							   PlanType == 2 ? MyConvert.ZHLC("夜班") :
							   MyConvert.ZHLC("异常");
					}
				}

				public string PlanRdsNo { get; set; }


				public DateTime ShowDate
				{
					get
					{
						return PlanBegin.Date;
					}
				}
				public DateTime PlanBegin { get; set; }

				public DateTime PlanEnd { get; set; }

				public TimeSpan WorkTime
				{
					get
					{
						return PlanEnd - PlanBegin;
					}
				}
				/// <summary>
				/// 是否大于6小时
				/// </summary>
				public bool isOver6Hour
				{
					get
					{
						return WorkTime.TotalHours > 6.100;
					}
				}

				public bool isSmall6Hour
				{
					get
					{
						return !isOver6Hour;
					}
				}

				public bool isRejectDay
				{
					get
					{
						return QualityYieldView < 95;
					}
				}

				public int DepartmentID { get; set; }

				public int DepartmentHrNumb { get; set; }

				public string DepartmentName { get; set; }

				public string ProcessCode { get; set; }

				public string ProcessName { get; set; }

				public string MachineCode { get; set; }

				public string MachineName { get; set; }

				public double MachineHrNumb { get; set; }

				public double PlanSheetNumb { get; set; }

				public double FinishSheetNumb { get; set; }
				/// <summary>
				/// 笔数达成率
				/// </summary>
				public double Yield1
				{
					get
					{
						if (PlanSheetNumb == 0) return 0;
						return FinishSheetNumb / PlanSheetNumb;
					}
				}

				public double PlanNumb { get; set; }

				public double FinishNumb { get; set; }

				public double PlanProdNumb { get; set; }

				public double FinishProdNumb { get; set; }
				/// <summary>
				/// 产量达成率
				/// </summary>
				public double Yield2
				{
					get
					{
						if (PlanProdNumb == 0) return 0;
						return FinishProdNumb / PlanProdNumb;
					}
				}

				public double Yield
				{
					get
					{
						return Yield1 * Yield2;
					}
				}
				/// <summary>
				/// 冒顶达成率，比105高。
				/// </summary>
				public bool isLargeThanTopPerformanceYield
				{
					get
					{
						return Yield > TopPerformanceYield;
					}
				}

				public bool isLargeTopPerformanceYield
				{
					get
					{
						return Yield >= TopPerformanceYield;
					}
				}

				/// <summary>
				/// 达成率显示值
				/// </summary>
				public double PerformanceView
				{
					get
					{
						return ((isLargeThanTopPerformanceYield ? TopPerformanceYield : Yield) * 100);
					}
				}

				/// <summary>
				/// 每天的完工数
				/// </summary>
				public double DailyFinishNumb { get; set; }
				/// <summary>
				/// 每天按照排程开机累计的判定不良品
				/// </summary>
				public double DailyRejectNumb { get; set; }

				public double QualityYield
				{
					get
					{
						if (DailyFinishNumb == 0 || DailyFinishNumb - DailyRejectNumb <= 0) return 0;
						return (DailyFinishNumb - DailyRejectNumb) / DailyFinishNumb;
					}
				}

				public double QualityYieldView
				{
					get
					{
						return QualityYield * 100;
					}
				}

				public double RejectTimeNumb { get; set; }

				public bool isNotInspect { get; set; }

				public double InspectScore { get; set; }

				public double InspectScoreView
				{
					get
					{
						return (InspectScore - RejectTimeNumb / 10) * 100;
					}
				}

				public int WeekName { get; set; }
				/// <summary>
				/// 当日综合得分
				/// </summary>
				public double Score
				{
					get
					{
						return (PerformanceView * 0.3) + (QualityYieldView * 0.3) + (InspectScoreView * 0.4);
					}
				}

			}

			public class MonthRate : MyBase.MyEnumerable<DateRateItem>
			{
				public MonthRate(DateTime dtBegin, DateTime dtEnd)
				{
					DateBegin = dtBegin;
					DateEnd = dtEnd;
					Load();
				}

				public MonthRate()
				{

				}

				public DateTime DateBegin { get; set; }

				public DateTime DateEnd { get; set; }

				public void Load()
				{
					string SQL = @"
Select a.*,InspectScore2 = isNull(InspectScore,-100),
	   WeekName=DatePart(Wk,PlanBegin),
	   p.name as DepartmentName,
	   m.name as MachineName,
	   pp.name as ProcessName
from [_PMC_ProdPlan_YieldRate] a Left Outer Join [pbDept] p On a.DeaprtmentID = p.[_ID] 
								 Left Outer Join [moMachine] m on a.MachineCode = m.code
								 Left Outer Join [moProcedure] pp on a.ProcessCode = pp.code 
 Where PlanBegin Between @DBegin And @DEnd 
";
					MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
					{
						new MyData.MyDataParameter("@DBegin",DateBegin, MyData.MyDataParameter.MyDataType.DateTime),
						new MyData.MyDataParameter("@DEnd",DateEnd, MyData.MyDataParameter.MyDataType.DateTime)
					};
					using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, mps))
					{
						if (md.MyRows.IsNotEmptySet())
						{
							var v = from a in md.MyRows
									select new DateRateItem(a);
							_data = v.ToList();
						}
					}
				}
			}

			public class MachineRateListItem
			{

				public MonthRate MonthList { get; set; }
				public string DepartmentName { get; set; }

				public string ProcessName { get; set; }

				public string MachineName { get; set; }
			}

			public class MachineRateList : MyBase.MyEnumerable<MachineRateListItem>
			{
				public MachineRateList(DateTime dtBegin, DateTime dtEnd)
				{
					DateBegin = dtBegin;
					DateEnd = dtEnd;
					Load();
				}

				public MachineRateList()
				{

				}

				public DateTime DateBegin { get; set; }

				public DateTime DateEnd { get; set; }

				public int DepartmentID { get; set; }


				public MonthRate AllRateList { get; set; }

				public void Load()
				{
					AllRateList = new MonthRate(DateBegin, DateEnd);
					var v = from a in Machines
							where a.Status < 2 && a.PlanStatus && a.ProcessName.IsNotEmpty() && a.ProcessPType > 0 && a.DepartmentID > 0
							orderby a.DepartmentFullSortID, a.ProcessCode, a.Code
							select a;
					foreach (var item in v)
					{
						if (item.ProcessCode != "9030" && item.ProcessCode != "9035")
						{
							MachineRateListItem xItem = new MachineRateListItem();
							xItem.MonthList = new MonthRate();
							xItem.DepartmentName = item.DepartmentName;
							xItem.ProcessName = item.ProcessName;
							xItem.MachineName = item.Name;
							var vx = AllRateList.Where(x => x.MachineCode == item.Code);
							if (vx.IsNotEmptySet()) xItem.MonthList.AddRange(vx.ToArray());
							_data.Add(xItem);
						}
					}
				}
			}
		}

		#endregion

		#region 发送看板

		void SendKanbanEmail_Performance(DateTime NowTime)
		{
			try
			{
				MyRecord.Say("-----------计算月达成率发送超过100的----------");
				MyRecord.Say("1.加载邮件体。");
				#region 表头
				string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{4}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {0:yy年MM月}高分达成率统计。</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; （说明：{0:yy/MM/dd} - {1:yy/MM/dd}之间生产看板达成率≥100%达到15天的机台。）</FONT></DIV>
{2}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{3:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");
				string gd = @"
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
	班别
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: transparent"" align=center >
	达成率达到100%次数
	</TD>
	</TR>
	{0}
</TBODY></TABLE></FONT>
</DIV>
";
				string sd = @"
<DIV><FONT size=5 face=PMingLiU color=#ff0080>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;本月没有达成率≥100%达到15天的机台。</FONT>
</DIV>
";
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
	</TR>
";
				#endregion
				if (NowTime > DateTime.MinValue && NowTime.Day == 1)
				{
					DateTime dtBegin = NowTime.AddMonths(-1).Date, dtEnd = NowTime.Date.AddMilliseconds(-10);
					MyRecord.Say(string.Format("3.读取数据，时间：{0}~{1}", dtBegin, dtEnd));
					ProductionScheduingRate.MachineRateList mrla = new ProductionScheduingRate.MachineRateList(dtBegin, dtEnd);
					MyRecord.Say("3.1 数据读取完毕。");
					var mrl = from a in mrla
							  where a.MonthList.IsNotEmptySet()
							  select a;
					string brs = string.Empty;
					int xRow = 0;
					foreach (var item in mrl)
					{
						// 白班的得分总计
						var vdSum = from a in item.MonthList
									where a.PlanType == 1
									select a;
						if (vdSum.IsNotEmptySet())
						{
							var vSumEffectiveDays = vdSum.Where(x => x.isLargeTopPerformanceYield);
							if (vSumEffectiveDays.Count() >= 15)
							{
								MyRecord.Say(string.Format("{0}，{1}，{2}，{3}，{4}", item.DepartmentName, item.ProcessName, item.MachineName, LCStr("白班"), vSumEffectiveDays.Count()));
								brs += string.Format(br, item.DepartmentName, item.ProcessName, item.MachineName, LCStr("白班"), vSumEffectiveDays.Count());
								xRow++;
							}
						}

						//夜班的得分总计
						var vnSum = from a in item.MonthList
									where a.PlanType == 2
									select a;
						if (vnSum.IsNotEmptySet())
						{
							var vSumEffectiveDays = vnSum.Where(x => x.isLargeTopPerformanceYield);
							if (vSumEffectiveDays.Count() >= 15)
							{
								MyRecord.Say(string.Format("{0}，{1}，{2}，{3}，{4}", item.DepartmentName, item.ProcessName, item.MachineName, LCStr("夜班"), vSumEffectiveDays.Count()));
								brs += string.Format(br, item.DepartmentName, item.ProcessName, item.MachineName, LCStr("夜班"), vSumEffectiveDays.Count());
								xRow++;
							}
						}
					}
					string gds = string.Format(gd, brs);
					if (xRow == 0) gds = sd;
					MyRecord.Say("创建SendMail。");
					MyBase.SendMail sm = new MyBase.SendMail();
					MyRecord.Say("加载邮件内容。");
					sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, dtBegin, dtEnd, gds, DateTime.Now, MyBase.CompanyTitle));
					sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yyyy年MM月}达成率高分统计表(测试)", dtBegin, MyBase.CompanyTitle));
					string MailTo = ConfigurationManager.AppSettings["Kanban_Performance_MailTo"], MailCC = ConfigurationManager.AppSettings["Kanban_Performance_MailCC"];
					//string MailTo = "my80@my.imedia.com.tw", MailCC = "";
					MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", MailTo, MailCC));
					sm.MailTo = MailTo;
					sm.MailCC = MailCC;
					//sm.MailTo = "my80@my.imedia.com.tw";
					MyRecord.Say("发送邮件。");
					sm.SendOut();
					MyRecord.Say("已经发送。");
				}
			}
			catch (Exception ex)
			{
				MyRecord.Say(ex);
			}
			MyRecord.Say("-------计算排程达成率完成------");
		}

		void SendKanbanEmail_Inspect(DateTime NowTime)
		{
			try
			{
				MyRecord.Say("-----------计算看板纪律检查超过2红6黄的----------");
				MyRecord.Say("1.加载邮件体。");
				#region 表头
				string body = MyConvert.ZH_TW(@"
<HTML>
<BODY style=""FONT-SIZE: 9pt; FONT-FAMILY: PMingLiU"" leftMargin=5 topMargin=5 bgColor=#ece4f3 #ffffff>
<DIV><FONT size=3 face=PMingLiU>{4}ERP系统提示您：</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU color=#FFFF0000>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; {0:yyyy年MM月}纪律稽核检查低分提醒。（测试，正式邮件会在每月1日8点发送）</FONT></DIV>
<DIV><FONT size=3 face=PMingLiU>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; （说明：{0:yy/MM/dd} - {1:yy/MM/dd}之间纪律稽核检查得分2红或6黄的机台，红：低于90分、黄：90~97分）</FONT></DIV>
{2}
<DIV><FONT face=PMingLiU><FONT size=2></FONT>&nbsp;</DIV>
<DIV><FONT color=#0000ff size=4 face=PMingLiU><STRONG>&nbsp;&nbsp;此郵件由ERP系統自動發送，请勿在此郵件上直接回復。</STRONG></FONT></DIV>
<DIV><FONT color=#800080 size=2><STRONG>&nbsp;&nbsp;&nbsp;</STRONG>
<FONT color=#000000 face=PMingLiU>{3:yy/MM/dd HH:mm}，由ERP系统伺服器自动发送。<BR>
&nbsp;&nbsp;&nbsp;&nbsp;如自動發送功能有問題或者格式内容修改建議，請MailTo:<A href=""mailto:my80@my.imedia.com.tw"">JOHN</A><BR>
</FONT></FONT></DIV></FONT></BODY></HTML>
");
				string gd = @"
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
	班别
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: #FFFCEF0D"" align=center >
	90~97分(黄)
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: #FFFC1C0D"" align=center >
	低于90分(红)
	</TD>
	</TR>
	{0}
</TBODY></TABLE></FONT>
</DIV>
";
				string sd = @"
<DIV><FONT size=5 face=PMingLiU color=#ff0080>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;没有。</FONT>
</DIV>
";
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
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: #FFFCEF0D"" align=center >
	{4}
	</TD>
	<TD class=xl63 style=""BORDER-TOP: windowtext 0.5pt solid; BORDER-RIGHT: windowtext 0.5pt solid; BORDER-BOTTOM: windowtext 0.5pt solid; BORDER-LEFT: windowtext 0.5pt solid; BACKGROUND-COLOR: #FFFC1C0D"" align=center >
	{5}
	</TD>
	</TR>
";
				#endregion
				if (NowTime > DateTime.MinValue && NowTime.Day == 1)
				{
					DateTime dtBegin = NowTime.AddMonths(-1).Date, dtEnd = NowTime.Date.AddMilliseconds(-10);
					MyRecord.Say(string.Format("3.读取数据，时间：{0}~{1}", dtBegin, dtEnd));
					ProductionScheduingRate.MachineRateList mrla = new ProductionScheduingRate.MachineRateList(dtBegin, dtEnd);
					MyRecord.Say("3.1 数据读取完毕。");
					var mrl = from a in mrla
							  where a.MonthList.IsNotEmptySet()
							  select a;
					string brs = string.Empty;
					int xRow = 0;
					foreach (var item in mrl)
					{
						// 白班的得分总计
						var vdSum = from a in item.MonthList
									where a.PlanType == 1
									select a;
						if (vdSum.IsNotEmptySet())
						{
							var vSumEffectiveDays = vdSum.Where(x => x.InspectScoreView <= 90);
							var vSumEffectiveDays2 = vdSum.Where(x => x.InspectScoreView > 90 && x.InspectScoreView <= 98);
							if (vSumEffectiveDays.Count() >= 2 || vSumEffectiveDays2.Count() >= 6)
							{
								MyRecord.Say(string.Format("{0}，{1}，{2}，{3}，红{4},黄{5}", item.DepartmentName, item.ProcessName, item.MachineName, LCStr("白班"), vSumEffectiveDays.Count(), vSumEffectiveDays2.Count()));
								brs += string.Format(br, item.DepartmentName, item.ProcessName, item.MachineName, LCStr("白班"), vSumEffectiveDays2.Count(), vSumEffectiveDays.Count());
								xRow++;
							}
						}

						//夜班的得分总计
						var vnSum = from a in item.MonthList
									where a.PlanType == 2
									select a;
						if (vnSum.IsNotEmptySet())
						{
							var vSumEffectiveDays = vnSum.Where(x => x.InspectScoreView <= 90);
							var vSumEffectiveDays2 = vnSum.Where(x => x.InspectScoreView > 90 && x.InspectScoreView <= 98);
							if (vSumEffectiveDays.Count() >= 2 || vSumEffectiveDays2.Count() >= 6)
							{
								MyRecord.Say(string.Format("{0}，{1}，{2}，{3}，红{4},黄{5}", item.DepartmentName, item.ProcessName, item.MachineName, LCStr("夜班"), vSumEffectiveDays.Count(), vSumEffectiveDays2.Count()));
								brs += string.Format(br, item.DepartmentName, item.ProcessName, item.MachineName, LCStr("夜班"), vSumEffectiveDays2.Count(), vSumEffectiveDays.Count());
								xRow++;
							}
						}
					}
					string gds = string.Format(gd, brs);
					if (xRow == 0) gds = sd;
					MyRecord.Say("创建SendMail。");
					MyBase.SendMail sm = new MyBase.SendMail();
					MyRecord.Say("加载邮件内容。");
					sm.MailBodyText = MyConvert.ZH_TW(string.Format(body, dtBegin, dtEnd, gds, DateTime.Now, MyBase.CompanyTitle));
					sm.Subject = MyConvert.ZH_TW(string.Format("{1}{0:yyyy年MM月}纪律稽核检查低分提醒(测试)", dtBegin, MyBase.CompanyTitle));
					string MailTo = ConfigurationManager.AppSettings["Kanban_Inspect_MailTo"], MailCC = ConfigurationManager.AppSettings["Kanban_Inspect_MailCC"];
					MyRecord.Say(string.Format("MailTO:{0}\nMailCC:{1}", MailTo, MailCC));
					sm.MailTo = MailTo;
					sm.MailCC = MailCC;
					//sm.MailTo = "my80@my.imedia.com.tw";
					MyRecord.Say("发送邮件。");
					sm.SendOut();
					MyRecord.Say("已经发送。");
				}
			}
			catch (Exception ex)
			{
				MyRecord.Say(ex);
			}
			MyRecord.Say("-------计算排程达成率完成------");
		}

		#endregion


		#region 机器机台类

		/// <summary>
		/// 所有机台存储器
		/// </summary>
		private static Machine _Machines = null;

		public static Machine Machines
		{
			get
			{
				if (_Machines == null)
					_Machines = new Machine();
				else if (_Machines.needReload)
					_Machines.reLoad();
				return _Machines;
			}
		}

		[Serializable]
		public class Machine : MachineCollection
		{
			public Machine()
				: base()
			{
				reLoad();
			}
			/// <summary> 重新加载
			///
			/// </summary>
			public void reLoad()
			{
				try
				{
					MyRecord.Say("----初始化机台基本资料-----");
					string sql = @"
Select procd as ProcCode,a.Code as MachineCode,a.Name as MachineName,
	   PrepTime,StdCapcity,a.remark,DepartmentID,a.[_id] as Machine_ID,
	   MachineStatus=Status,StaticMeasurement,
	   a.Remark,Colors,GeneralLossFix,GeneralLossRate,PlanMeasurement,FinishMeasurement,StdPersonNumber,PlanStatus,TypeName,
	   WithoutLevel=isNull(WithoutLevel,0),MustHaveRejectNumber=isNull(MustHaveRejectNumber,0),AutoDuplexPrint=isNull(AutoDuplexPrint,0),
	   b.name as ProcessName,c.name as DepartmentName,b.PType as ProcessPType
	   From moMachine a Left Outer Join moProcedure b On a.procd = b.code 
						Left Outer Join pbDept c On a.DepartmentID = c.[_ID]
		Where isNull(Status,0)<2 Order by Procd,MachineName";
					using (MyData.MyDataTable md = new MyData.MyDataTable(sql))
					{
						MyRecord.Say(string.Format("读取了：{0}条机台信息",md.MyRows.Count));
						reLoad(md.MyRows);
						_LastChangeTime = TestLastChangeTime = LastLoadTime = DateTime.Now;
						LoadTimes++;
					}
				}
				catch (Exception ex)
				{
					MyRecord.Say(ex);
				}
				MyRecord.Say("----初始化机台基本资料，完毕。-----");
			}



			/// <summary>
			/// 最后加载时间
			/// </summary>
			public DateTime LastLoadTime { get; private set; }
			/// <summary>
			/// 加载的次数
			/// </summary>
			public int LoadTimes { get; private set; }

			private DateTime _LastChangeTime = DateTime.Now, TestLastChangeTime = DateTime.Now;
			/// <summary>
			/// 最后修改日期
			/// </summary>
			public DateTime LastChangeTime
			{
				get
				{
					if ((DateTime.Now - TestLastChangeTime).TotalMinutes > 5)
					{
						string SQL = "Select Max([Date]) from [_PMC_Machine_StatusRecorder]";
						object v = MyData.MyCommand.ExecuteScalar(SQL);
						DateTime oValue;
						if (v.IsNotNull())
						{
							if (DateTime.TryParse(v.ToString(), out oValue))
							{
								_LastChangeTime = oValue;
							}
							else
							{
								_LastChangeTime = DateTime.MinValue;
							}
						}
						TestLastChangeTime = DateTime.Now;
					}
					return _LastChangeTime;
				}
			}

			public bool needReload
			{
				get
				{
					return (DateTime.Now - LastLoadTime).TotalMinutes > 75 || (LastChangeTime - LastLoadTime).Minutes > 0;
				}
			}
			/// <summary>
			/// 空操作，为了检查更新
			/// </summary>
			public void DoNop()
			{

			}
		}

		/// <summary>
		/// 机台集合类
		/// </summary>
		[Serializable]
		public class MachineCollection : MyBase.MyEnumerable<MachineItem>
		{
			#region 构造函数

			protected void reLoad(MyData.MyDataRowCollection MachineData)
			{
				if (MachineData != null && MachineData.Count > 0)
				{
					var v = from a in MachineData
							select new MachineItem(a.Value("MachineCode"))
							{
								Name = a.Value("MachineName"),
								ProcessCode = a.Value("ProcCode"),
								ProcessName = a.Value("ProcessName"),
								DepartmentID = a.IntValue("DepartmentID"),
								DepartmentName = a.Value("DepartmentName"),
								DepartmentFullSortID = a.Value("FullSortID"),
								ProcessPType = a.IntValue("ProcessPType"),
								StdPerpareTime = a.IntValue("Preptime"),
								StdCapacity = a.Value<double>("StdCapcity"),
								StdPersonNumber = a.IntValue("StdPersonNumber"),
								ID = a.IntValue("Machine_ID"),
								Status = a.IntValue("MachineStatus"),
								GeneralLossFix = a.Value<float>("GeneralLossFix"),
								GeneralLossRate = a.Value<double>("GeneralLossRate"),
								PlanMeasurement = a.IntValue("PlanMeasurement"),
								FinishMeasurement = a.IntValue("FinishMeasurement"),
								StaticMeasurement = a.IntValue("StaticMeasurement"),
								Memo = a.Value("remark"),
								Colors = a.IntValue("Colors"),
								WithoutLevel = a.BooleanValue("WithoutLevel"),
								MustHaveRejectNumber = a.BooleanValue("MustHaveRejectNumber"),
								PlanStatus = a.BooleanValue("PlanStatus"),
								AutoDuplexPrint = a.BooleanValue("AutoDuplexPrint"),
								TypeName = a.Value("TypeName")
							};
					reLoad(v);
				}
			}

			protected void reLoad(IEnumerable<MachineItem> ieMachineData)
			{
				_data = null;
				if (ieMachineData != null && ieMachineData.Count() > 0)
				{
					_data = ieMachineData.ToList();
				}
			}

			public MachineCollection()
			{
			}

			public MachineCollection(MyData.MyDataRowCollection MachineData)
			{
				reLoad(MachineData);
			}

			public MachineCollection(IEnumerable<MachineItem> ieMachineData)
			{
				reLoad(ieMachineData);
			}

			#endregion 构造函数

			/// <summary>
			/// 机台是否存在
			/// </summary>
			/// <param Name="MachineCode"></param>
			/// <returns></returns>
			public bool Contains(string MachineCode)
			{
				var v = from a in _data
						where a.Code == MachineCode
						select a;
				return (v.IsNotNull() && v.Count() > 0);

			}

			/// <summary>
			/// 获取某一工序下所有机台。
			/// </summary>
			/// <param Name="processcode"></param>
			/// <returns></returns>
			public MachineCollection getMachinesByProcess(string processcode)
			{
				var v = from a in _data
						where a.ProcessCode == processcode
						select a;
				return new MachineCollection(v);
			}

			/// <summary>
			/// 获取某一部门下所有机台
			/// </summary>
			/// <param Name="deptID"></param>
			/// <returns></returns>
			public MachineCollection getMachinesByDepartment(int deptID)
			{
				var v = from a in _data
						where a.DepartmentID == deptID
						select a;
				return new MachineCollection(v);
			}

			#region 索引器

			public MachineItem this[string keyMachineCode]
			{
				get
				{
					var v = from a in _data
							where a.Code == keyMachineCode
							select a;
					if (v.IsNotNull() && v.Count() > 0)
						return v.FirstOrDefault();
					else
						return null;
				}
				set
				{
					var v = from a in _data
							where a.Code == keyMachineCode
							select a;
					if (v.IsNotNull() && v.Count() > 0)
					{
						MachineItem xItem = v.FirstOrDefault();
						xItem = value;
					}
					else
					{
						_data.Add(value);
					}
				}
			}

			#endregion 索引器


		}

		/// <summary>
		/// 机台类
		/// </summary>
		[Serializable]
		public class MachineItem
		{
			#region 存储器

			/// <summary>
			/// 机台编号
			/// </summary>
			public string Code { get; set; }

			/// <summary>
			/// 机台名称
			/// </summary>
			public string Name { get; set; }

			/// <summary>
			/// 标准调机时间
			/// </summary>
			public int StdPerpareTime { get; set; }

			/// <summary>
			/// 标准产能
			/// </summary>
			public double StdCapacity { get; set; }

			/// <summary>
			/// 标准人数
			/// </summary>
			public int StdPersonNumber { get; set; }


			/// <summary>
			/// 标准固定放损
			/// </summary>
			public float GeneralLossFix { get; set; }

			/// <summary>
			/// 标准放损比例
			/// </summary>
			public double GeneralLossRate { get; set; }

			/// <summary>
			/// 部门ID
			/// </summary>
			public int DepartmentID { get; set; }

			public string DepartmentName { get; set; }

			public string DepartmentFullSortID { get; set; }
			/// <summary>
			/// 工序编号
			/// </summary>
			public string ProcessCode { get; set; }

			public string ProcessName { get; set; }

			public int ProcessPType { get; set; }

			/// <summary>
			/// 排期产能计量标准
			/// 0-按产品数计量、1-按车头数计量
			/// </summary>
			public int PlanMeasurement { get; set; }

			/// <summary>
			/// 完工单和达成率输入计量标准
			/// 0-按产品数计量、1-按车头数计量
			/// </summary>
			public int FinishMeasurement { get; set; }

			/// <summary>
			/// 完工单和达成率计算计量标准
			/// 0-只计算良品数、1-良品数和不良品数一起计算
			/// </summary>
			public int StaticMeasurement { get; set; }


			/// <summary>
			/// ID
			/// </summary>
			public int ID { get; set; }

			/// <summary>
			/// 状态
			/// </summary>
			public int Status { get; set; }
			/// <summary>
			/// 是否进入排程
			/// </summary>
			public bool PlanStatus { get; set; }

			/// <summary>
			/// 说明
			/// </summary>
			public string Memo { get; set; }

			/// <summary>
			/// 颜色数
			/// </summary>
			public int Colors { get; set; }
			/// <summary>
			/// 不用设置机长等级，默认等级是11
			/// </summary>
			public bool WithoutLevel { get; set; }
			/// <summary>
			/// 完工都必须输入不合格品数量。
			/// </summary>
			public bool MustHaveRejectNumber { get; set; }
			/// <summary>
			/// 是否自动双面（双翻）
			/// </summary>
			public bool AutoDuplexPrint { get; set; }
			/// <summary>
			/// 归类名称
			/// </summary>
			public string TypeName { get; set; }

			#endregion 存储器

			#region 构造器

			public MachineItem(string code)
			{
				Code = code;
			}

			#endregion 构造器


		}

		#endregion 机器机台类
	}
}
