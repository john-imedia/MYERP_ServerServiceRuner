using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.ClearScript.V8;
using MYERP_ServerServiceRuner.Base;

namespace MYERP_ServerServiceRuner
{
    public partial class MainService
    {
        #region 估价
        [Serializable]
        public class EstimateNote
        {
            public EstimateNote()
            {
                Products = new EstimateCollection();
                isNew = true;
            }

            public EstimateNote(Guid xid)
            {
                CGUID = xid;
            }

            public string RdsNo { get; set; }

            public string Inputer { get; set; }

            public string Checker { get; set; }

            public DateTime InputDate { get; set; }

            public DateTime CheckDate { get; set; }

            public int Status { get; set; }

            public Guid CGUID { get; set; }

            public int id { get; set; }

            public bool isNew { get; set; }

            public EstimateCollection Products { get; set; }

            public string CustomerCode
            {
                get
                {
                    return Customer.IsNotNull() ? Customer.Code : string.Empty;
                }
                set
                {
                    if (Customers.Contains(value))
                        Customer = MyServiceLogic.Customers[value];
                    else
                        Customer = null;
                }
            }

            public MyServiceLogic.CustomerItem Customer { get; set; }


            string _MoneyTypeCode = null;

            public string MoneyTypeCode
            {
                get
                {
                    if (_MoneyTypeCode.IsEmpty())
                        return MoneyTypeCode_Default;
                    else
                        return _MoneyTypeCode;
                }
                set
                {
                    _MoneyTypeCode = value;
                }
            }

            public double Numb1 { get; set; }

            public double Numb2 { get; set; }

            public double Numb3 { get; set; }

            public double Price1 { get; set; }

            public double Price2 { get; set; }

            public double Price3 { get; set; }

            public double Amount1 { get; set; }

            public double Amount2 { get; set; }

            public double Amount3 { get; set; }

            /// <summary>
            /// KIT 的Code
            /// </summary>
            public string Code { get; set; }
            /// <summary>
            /// KIT的Name
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 是不是按照KIT报价
            /// </summary>
            public bool isKIT { get; set; }


            public void LoadKITBom()
            {
                if (!isKIT) return;
                if (Code.IsEmpty()) return;
                if (!MyServiceLogic.Products.Contains(this.Code)) return;
                ProductItem thisProduct = MyServiceLogic.Products[this.Code];
                Name = thisProduct.Name;
                CustomerCode = thisProduct.CustCode;
                wp.TitleWord = "加载KIT印件BOM的半成品。";
                wp.DescriptionWord = "请稍等。";
                thisProduct.BOM.reLoadDetial();

                if (thisProduct.BOM.Materials.IsNotNull() && thisProduct.BOM.Materials.Count > 0)
                {
                    Products = new EstimateCollection();

                    EstimateItem EMI = new EstimateItem(this);
                    Products.Add(EMI);
                    EMI.Code = this.Code;
                    EMI.LoadFromBom(wp);
                    EMI.Remark = MySimplyBasic.MyConvert.ZHLC("KIT产品加工费");
                    int emIndex = 0;
                    do
                    {
                        EstimateMaterialItem emItem = EMI.Material[emIndex];
                        if (emItem.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Material && emItem.MaterialItem.Type.Type == MaterialType.MaterialTypeEnum.Product)
                        {
                            EMI.Material.RemoveItem(emItem);
                        }
                        else
                        {
                            emIndex++;
                        }
                    } while (emIndex < EMI.Material.Count());

                    var v = from a in thisProduct.BOM.Materials
                            orderby a.Sort_ID
                            select a;

                    foreach (var item in v)
                    {
                        if (item.MaterialType == BomMaterial.BomMaterialTypeEnum.BomProduct)
                        {
                            EstimateItem EI = new EstimateItem(this);
                            Products.Add(EI);
                            wp.TitleWord = string.Format("正在加载，{0}", item.Code);
                            Application.DoEvents();
                            EI.Code = item.Code;
                            EI.LoadFromBom(wp);
                        }
                    }
                }
            }

            public void Calculate()
            {
                foreach (var item in Products)
                {
                    item.Calculate();
                }
            }
        }
        [Serializable]
        public class EstimateCollection : MyBase.MyEnumerable<EstimateItem>
        {

        }
        [Serializable]
        public class EstimateItem
        {
            public delegate void FillGridHanler(C1.Win.C1FlexGrid.Row r, EstimateItem x);

            public FillGridHanler FillGrid;
            /// <summary>
            /// 用用从BOM加载
            /// </summary>
            /// <param name="parent"></param>
            public EstimateItem(EstimateNote parent)
            {
                isNew = true;
                Parent = parent;
                var v = from a in parent.Products
                        select a.id;
                if (v.IsNotNull() && v.Count() > 0)
                {
                    id = v.Max() + 1;
                }
                else
                {
                    id = 1;
                }

                Material = new EstimateMaterialCollection();
                Process = new EstimateProcessCollection();
                Charge = new EstimateChargeCollection();
            }

            public EstimateItem(EstimateNote parent, C1.Win.C1FlexGrid.Row GridRow)
            {
                isNew = true;
                CurrentGridRow = GridRow;
                GridRow.UserData = this;
                Parent = parent;
                var v = from a in parent.Products
                        select a.id;
                if (v.IsNotNull() && v.Count() > 0)
                {
                    id = v.Max() + 1;
                }
                else
                {
                    id = 1;
                }

                Material = new EstimateMaterialCollection();
                Process = new EstimateProcessCollection();
                Charge = new EstimateChargeCollection();
            }

            public EstimateItem(EstimateNote parent, MyData.MyDataRow mr)
            {
                Parent = parent;
                isNew = false;
                id = Convert.ToInt32(mr["PID"]);
                isBook = Convert.ToBoolean(mr["isBook"]);
                Name = Convert.ToString(mr["Name"]);
                Code = Convert.ToString(mr["Code"]);
                TypeCode = Convert.ToString(mr["TypeCode"]);
                EstimateRdsNo = Convert.ToString(mr["EstimateRdsNo"]);
                Remark = Convert.ToString(mr["Remark"]);
                MaterialRemark = Convert.ToString(mr["MaterialRemark"]);
                Length = Convert.ToDouble(mr["Length"]);
                Width = Convert.ToDouble(mr["Width"]);
                Height = Convert.ToDouble(mr["Height"]);
                UnflodLength = Convert.ToDouble(mr["UnflodLength"]);
                UnflodWidth = Convert.ToDouble(mr["UnflodWidth"]);

                Material = new EstimateMaterialCollection();
                Process = new EstimateProcessCollection();
                Charge = new EstimateChargeCollection();
            }

            public EstimateNote Parent { get; set; }

            public bool isNew { get; set; }

            public int id { get; set; }

            public bool isBook { get; set; }

            public C1.Win.C1FlexGrid.Row CurrentGridRow { get; set; }
            /// <summary>
            /// 料号
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 产品编号，用于从BOM加载
            /// </summary>
            public string Code { get; set; }
            /// <summary>
            /// 产品分类编号。
            /// </summary>
            public string TypeCode { get; set; }

            public MaterialTypeItem Type
            {
                get
                {
                    return MaterialTypes[TypeCode];
                }
            }

            public string EstimateRdsNo { get; set; }

            public string Remark { get; set; }

            public string MaterialRemark { get; set; }

            public string Size
            {
                get
                {
                    if (Length + Width + Height == 0)
                        return string.Empty;
                    else if (Length * Width * Height == 0)
                    {
                        if (Length == 0)
                        {
                            return string.Format("{0:0.##}x{1:0.##}", Width, Height);
                        }
                        else
                        {
                            return string.Format("{0:0.##}x{1:0.##}", Length, Width == 0 ? Height : Width);
                        }
                    }
                    else
                        return string.Format("{0:0.##}x{1:0.##}x{2:0.##}", Length, Width, Height);
                }
                set
                {
                    string uSize = value.Trim();
                    Length = Width = Height = 0;
                    Regex rgx = new Regex(@"^([\d\.]+)(?=[xX*])");
                    string SL = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    rgx = new Regex(@"(?<=[xX*])([\d\.]+)(?=[xX*])");
                    string SW = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    rgx = new Regex(@"(?<=[xX*])([\d\.]+)$");
                    string SH = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    double l = 0, w = 0, h = 0;
                    if (double.TryParse(SL, out l)) Length = l;
                    if (double.TryParse(SW, out w)) Width = w;
                    if (double.TryParse(SH, out h))
                    {
                        if (w == 0) Width = h;
                        else Height = h;
                    }
                }
            }

            /// <summary>
            /// 成品长
            /// </summary>
            public double Length { get; set; }

            /// <summary>
            /// 成品宽
            /// </summary>
            public double Width { get; set; }

            /// <summary>
            /// 成品高
            /// </summary>
            public double Height { get; set; }

            public string UnflodSize
            {
                get
                {
                    return string.Format("{0:0.##}x{1:0.##}", UnflodLength, UnflodWidth);
                }
                set
                {
                    string uSize = value.Trim();
                    UnflodLength = UnflodWidth = 0;
                    Regex rgx = new Regex(@"^([\d\.]+)(?=[xX*])");
                    string uFL = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    rgx = new Regex(@"(?<=[xX*])([\d\.]+)$");
                    string uFW = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    double ufl = 0, ufw = 0;
                    if (double.TryParse(uFL, out ufl)) UnflodLength = ufl;
                    if (double.TryParse(uFW, out ufw)) UnflodWidth = ufw;
                }
            }

            public double UnflodLength { get; set; }

            public double UnflodWidth { get; set; }


            public double PaperPurchasePrice
            {
                get
                {
                    var v = from a in Material
                            where a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().Price;
                    }
                    return 0;
                }
            }

            public double PaperGW
            {
                get
                {
                    var v = from a in Material
                            where a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().gw;
                    }
                    return 0;
                }
            }

            public string PaperSize
            {
                get
                {
                    var v = from a in Material
                            where a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().MaterialSize;
                    }
                    return string.Empty;
                }
            }

            public double PaperColNumb
            {
                get
                {
                    var v = from a in Material
                            where a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().ColNumb;
                    }
                    return 1;
                }
            }

            public double PaperCutNumb
            {
                get
                {
                    var v = from a in Material
                            where a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().CutNumb;
                    }
                    return 1;
                }
            }

            public double PaperPrice
            {
                get
                {
                    if (Numb > 0)
                    {
                        double iAmount = 0;
                        var v = from a in Material
                                where a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                                select a;
                        foreach (var item in v)
                        {
                            iAmount += item.Amount;
                        }
                        return iAmount / Numb;
                    }
                    return 0;
                }
            }

            public double WavePaperPrice
            {
                get
                {
                    if (Numb > 0)
                    {
                        double iAmount = 0;
                        var v = from a in Material
                                where a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_WavePaper
                                select a;
                        foreach (var item in v)
                        {
                            iAmount += item.Amount;
                        }
                        return iAmount / Numb;
                    }
                    return 0;
                }
            }

            public double MaterialPrice
            {
                get
                {
                    if (Numb > 0)
                    {
                        double iAmount = 0;
                        var v = from a in Material
                                where a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Material
                                select a;
                        foreach (var item in v)
                        {
                            iAmount += item.Amount;
                        }
                        return iAmount / Numb;
                    }
                    return 0;
                }
            }

            public double PrintPrice
            {
                get
                {
                    if (Numb > 0)
                    {
                        double iAmount = 0;
                        var v = from a in Process
                                where a.isPrint
                                select a;
                        foreach (var item in v)
                        {
                            iAmount += item.Amount;
                        }
                        return iAmount / Numb;
                    }
                    return 0;
                }
            }

            public double PostprocessPrice
            {
                get
                {
                    if (Numb > 0)
                    {
                        double iAmount = 0;
                        var v = from a in Process
                                where a.isNotPrint
                                select a;
                        foreach (var item in v)
                        {
                            iAmount += item.Amount;
                        }
                        return iAmount / Numb;
                    }
                    return 0;
                }
            }

            public double ChargePrice
            {
                get
                {
                    if (Numb > 0)
                    {
                        double iAmount = 0;
                        var v = from a in Charge
                                select a;
                        foreach (var item in v)
                        {
                            iAmount += item.Amount;
                        }
                        return iAmount / Numb;
                    }
                    return 0;
                }
            }

            public double Price1
            {
                get
                {
                    Numb = Numb1;
                    SetNumber();
                    return GotPrice();
                }
            }

            public double Price2
            {
                get
                {
                    Numb = Numb2;
                    SetNumber();
                    return GotPrice();
                }
            }

            public double Price3
            {
                get
                {
                    Numb = Numb3;
                    SetNumber();
                    return GotPrice();
                }
            }
            /// <summary>
            /// 当前计算的数量
            /// </summary>
            public double Numb { get; set; }

            public double Numb1
            {
                get
                {
                    if (Parent.IsNotNull())
                        return Parent.Numb1;
                    else
                        return 0;
                }
                set
                {
                    Numb = value;
                    if (Parent.IsNotNull()) Parent.Numb1 = value;
                }
            }
            public double Numb2
            {
                get
                {
                    if (Parent.IsNotNull())
                        return Parent.Numb2;
                    else
                        return 0;
                }
                set
                {
                    Numb = value;
                    if (Parent.IsNotNull()) Parent.Numb2 = value;
                }
            }

            public double Numb3
            {
                get
                {
                    if (Parent.IsNotNull())
                        return Parent.Numb3;
                    else
                        return 0;
                }
                set
                {
                    Numb = value;
                    if (Parent.IsNotNull()) Parent.Numb3 = value;
                }
            }
            /// <summary>
            /// 绑定到表格，并且会自动计算
            /// </summary>
            public void Calculate()
            {
                if (FillGrid != null && CurrentGridRow != null) FillGrid(CurrentGridRow, this);
            }

            public void LoadPrice()
            {
                if (Process.IsNotNull() && Process.Count() > 0)
                {
                    Process.ForEach((EstimateProcessItem x) =>
                    {
                        x.LoadPrice();
                    });
                }
                if (Material.IsNotNull() && Material.Count() > 0)
                {
                    Material.ForEach((EstimateMaterialItem x) =>
                    {
                        x.LoadPrice();
                    });
                }
            }

            public void SetNumber()
            {
                if (Process.IsNotNull() && Process.Count() > 0)
                {
                    Process.ForEach((EstimateProcessItem x) =>
                    {
                        x.ProductNumb = this.Numb;
                    });
                }
                if (Material.IsNotNull() && Material.Count() > 0)
                {
                    Material.ForEach((EstimateMaterialItem x) =>
                    {
                        x.ProductNumb = this.Numb;
                    });
                }
                if (Charge.IsNotNull() && Charge.Count() > 0)
                {
                    Charge.ForEach((EstimateChargeItem x) =>
                    {
                        x.ProductNumb = this.Numb;
                    });
                }
            }

            public double GotPrice()
            {
                if (this.Numb > 0)
                {
                    double iAmount = 0;
                    if (Process.IsNotNull() && Process.Count() > 0)
                    {
                        var vProcess = from a in Process
                                       orderby a.Sort_ID descending
                                       select a;
                        vProcess.ForEach((EstimateProcessItem x) => { iAmount += x.Amount; });
                    }
                    if (Material.IsNotNull() && Material.Count() > 0)
                    {
                        Material.ForEach((EstimateMaterialItem x) =>
                        {
                            iAmount += x.Amount;
                        });
                    }
                    if (Charge.IsNotNull() && Charge.Count() > 0)
                    {
                        Charge.ForEach((EstimateChargeItem x) =>
                        {
                            iAmount += x.Amount;
                        });
                    }
                    return iAmount / this.Numb;
                }
                else
                {
                    return 0;
                }
            }

            /// <summary>
            /// 工序工价
            /// </summary>
            public EstimateProcessCollection Process { get; set; }
            /// <summary>
            /// 报价原材料、辅材料
            /// </summary>
            public EstimateMaterialCollection Material { get; set; }
            /// <summary>
            /// 费用
            /// </summary>
            public EstimateChargeCollection Charge { get; set; }


            public void LoadFromBom()
            {
                if (Products.Contains(this.Code))
                {
                    MyRecord.Say(string.Format("1.获取产品信息{0}。", this.Code));
                    ProductItem p = Products[this.Code];
                    this.Name = p.Name;
                    this.Size = p.Size;
                    this.UnflodSize = p.UnfoldSize;
                    this.isBook = p.MutiplePart;
                    this.TypeCode = p.TypeCode;
                    MyRecord.Say("2.加载BOM。");
                    Bom b = p.BOM;
                    b.reLoadDetial();
                    MyRecord.Say("3.加载工序信息。");
                    this.Process.Clear();
                    if (b.Processes.IsNotNull() && b.Processes.Count > 0)
                    {
                        var vProcess = from a in b.Processes
                                       orderby a.Sort_ID
                                       select new EstimateProcessItem(this)
                                       {
                                           ProcessCode = a.Process.IsNotNull() ? a.Process.Code : null,
                                           MachineCode = a.Machine.IsNotNull() ? a.Machine.Code : null,
                                           PlateNumb = a.PlateNumber,
                                           ColorNumb1 = a.CA1 + a.CB1,
                                           ColorNumb2 = a.CA2 + a.CB2,
                                           ColorNumb3 = GetGlazingNumber(string.Format("{0}/{1}", a.ProcMemo, a.OtherMemo)),
                                           StdCapcityRate = a.ProductionCoefficient,
                                           StdLossTimeRate = a.CleanMachineCoefficient,
                                           PartID = a.PartID,
                                           ColNumb = a.ColNumb,
                                           ProduceRemark = string.Format("{0}/{1}", a.ProcMemo, a.OtherMemo),
                                           id = a.ID
                                       };
                        this.Process.AddRange(vProcess.ToArray());
                        //带入单价和版数（版数在纸张里）
                        int iIndex = 0;
                        foreach (EstimateProcessItem item in this.Process)
                        {
                            iIndex++;
                            if (wp.IsNotNull()) wp.DescriptionWord = string.Format("3.{0}.{1}获取价格。", iIndex, item.ProcessName);
                            if (item.Process.isPrint)
                            {
                                var vPlate = from a in b.Materials
                                             where a.PartID == item.PartID && a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper && a.ProcCode == item.ProcessCode
                                             orderby a.Sort_ID
                                             select a;
                                if (vPlate.IsNotNull() && vPlate.Count() > 0)
                                {
                                    item.PlateNumb = vPlate.FirstOrDefault().PlateNumber;
                                }
                                if (item.Machine.IsNotNull() && item.Machine.AutoDuplexPrint && item.PlateNumb == 1)
                                {
                                    item.PlateType = MyConvert.ZHLC("自翻版");
                                }
                                else
                                {
                                    item.PlateType = MyConvert.ZHLC("正反版");
                                }
                            }
                        }
                    }
                    MyRecord.Say("4.加载物料信息。");
                    this.Material.Clear();
                    if (b.Materials.IsNotNull() && b.Materials.Count > 0)
                    {
                        var vMaterial = from a in b.Materials
                                        orderby a.Sort_ID
                                        select new EstimateMaterialItem(this,
                                            a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper ? EstimateMaterialItem.MaterialTypeEnum.Type_Paper :
                                            a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomWavePaper ? EstimateMaterialItem.MaterialTypeEnum.Type_WavePaper :
                                            EstimateMaterialItem.MaterialTypeEnum.Type_Material)
                                        {
                                            MaterialCode = a.Code,
                                            ColNumb = a.ColNumb,
                                            CutNumb = a.CutNumb,
                                            UnitNumb = a.UnitNumber,
                                            PartID = a.PartID,
                                            Length = a.Length,
                                            Width = a.Width,
                                            id = a.ID
                                        };
                        this.Material.AddRange(vMaterial.ToArray());
                    }

                    LoadPrice();
                }
            }


            double GetGlazingNumber(string MemoText)
            {
                string hzMemoText = MyConvert.LCZH(MemoText),
                       dblKeyWord = MyConvert.LCZH("双面"),
                       glazingKeyWord1 = MyConvert.LCZH("油"),
                       glazingKeyWord2 = MyConvert.LCZH("光");
                string rgxGlazingStr = string.Format("[{0},{1}]", glazingKeyWord1, glazingKeyWord2);
                int numb = 0;
                Regex rgx1 = new Regex(rgxGlazingStr);
                if (rgx1.IsMatch(hzMemoText))
                {
                    numb = 1;
                    if (hzMemoText.Contains(dblKeyWord))
                    {
                        numb = 2;
                    }
                }
                return numb;
            }
        }
        [Serializable]
        public class EstimateProcessCollection : MyBase.MyEnumerable<EstimateProcessItem>
        {
            public EstimateProcessCollection()
            {

            }

        }
        [Serializable]
        public class EstimateProcessItem
        {
            /// <summary>
            /// 用于BOM生成
            /// </summary>
            /// <param name="parents"></param>
            protected internal EstimateProcessItem(EstimateItem parents)
            {
                Parent = parents;
                isNew = true;
                var v = from a in parents.Process
                        select a.id;
                if (v.IsNotNull() && v.Count() > 0)
                {
                    id = v.Max() + 1;
                }
                //else
                //{
                //    id = 1;
                //}
            }
            /// <summary>
            /// 从BOM来的加工说明。
            /// </summary>
            public string ProduceRemark { get; set; }

            /// <summary>
            /// 加载默认单价
            /// </summary>
            public void LoadPrice()
            {
                string thisMoneyTypeCode = Parent.Parent.MoneyTypeCode;
                MoneyTypeItem mti = MoneyType[thisMoneyTypeCode];
                if (isPrint)
                {
                    IEnumerable<EstimateProcessPriceItem> v;
                    if (PaperGW < 300)
                    {
                        v = from a in EstimateProcessPriceList
                            where a.MachineType == (this.Machine.Colors > 2 ? 1 : 0) && a.ProcessCode == this.ProcessCode
                            orderby a.id
                            select a;
                    }
                    else
                    {
                        v = from a in EstimateProcessPriceList
                            where a.MachineType == -1 && a.ProcessCode == this.ProcessCode
                            orderby a.id
                            select a;
                    }
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        EstimateProcessPriceItem fProcessItem = v.FirstOrDefault();
                        Price1 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price1) : 0;
                        Price2 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price2) : 0;
                        Price3 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price3) : 0;
                        StartPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.StartPrice) : 0;
                        StartNumber = fProcessItem.StartNumber;
                        FixPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.FixPrice) : 0;
                        LossRatio = fProcessItem.LossRatio;
                        Name = fProcessItem.Name;
                    }
                }
                else if (isDiecut)
                {
                    var v = from a in EstimateProcessPriceList
                            where a.MachineCode == this.MachineCode && a.ProcessCode == this.ProcessCode
                            orderby a.id
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        EstimateProcessPriceItem fProcessItem = v.FirstOrDefault();
                        FormulaString = fProcessItem.FormulaString2.IsNotEmpty() ? fProcessItem.FormulaString2 : fProcessItem.FormulaString;
                        FormulaView = fProcessItem.FormulaView2.IsNotEmpty() ? fProcessItem.FormulaView2 : fProcessItem.FormulaView;
                        UsePersonFormula = fProcessItem.UsePersonFormula;
                        LossRatio = fProcessItem.LossRatio;
                        Unit = fProcessItem.Unit;

                        Price1 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price1) : 0;
                        Price2 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price2) : 0;
                        Price3 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price3) : 0;
                        StartPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.StartPrice) : 0;
                        StartNumber = fProcessItem.StartNumber;
                        FixPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.FixPrice) : 0;
                        LossRatio = fProcessItem.LossRatio;
                        Name = fProcessItem.Name;
                    }
                    else
                    {
                        double PrintSize = PrintLength * PrintWidth;
                        var vArea = from a in EstimateProcessPriceList
                                    where a.MachineCode.IsEmpty() && a.ProcessCode == this.ProcessCode && a.Type2 == 1 && a.BeginNumb < PrintSize
                                    orderby a.id descending
                                    select a;

                        var vColNumber = from a in EstimateProcessPriceList
                                         where a.MachineCode.IsEmpty() && a.ProcessCode == this.ProcessCode && a.Type2 == 2 && a.BeginNumb < ColNumb
                                         orderby a.id descending
                                         select a;

                        EstimateProcessPriceItem fProcessItem = null;
                        if (vArea.IsNotNull() && vArea.Count() > 0 && vColNumber.IsNotNull() && vColNumber.Count() > 0)
                        {
                            EstimateProcessPriceItem fProcessItem1 = vArea.FirstOrDefault(), fProcessItem2 = vColNumber.FirstOrDefault();
                            if (fProcessItem1.Price1 > fProcessItem2.Price1)
                                fProcessItem = fProcessItem1;
                            else
                                fProcessItem = fProcessItem2;
                        }
                        else if (vArea.IsNotNull() && vArea.Count() > 0)
                        {
                            fProcessItem = vArea.FirstOrDefault();
                        }
                        else if (vColNumber.IsNotNull() && vColNumber.Count() > 0)
                        {
                            fProcessItem = vColNumber.FirstOrDefault();
                        }
                        if (fProcessItem.IsNotNull())
                        {
                            FormulaString = fProcessItem.FormulaString2.IsNotEmpty() ? fProcessItem.FormulaString2 : fProcessItem.FormulaString;
                            FormulaView = fProcessItem.FormulaView2.IsNotEmpty() ? fProcessItem.FormulaView2 : fProcessItem.FormulaView;
                            UsePersonFormula = fProcessItem.UsePersonFormula;
                            LossRatio = fProcessItem.LossRatio;
                            Unit = fProcessItem.Unit;

                            Price1 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price1) : 0;
                            Price2 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price2) : 0;
                            Price3 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price3) : 0;
                            StartPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.StartPrice) : 0;
                            StartNumber = fProcessItem.StartNumber;
                            FixPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.FixPrice) : 0;
                            LossRatio = fProcessItem.LossRatio;
                            Name = fProcessItem.Name;
                        }
                    }
                }
                else if (isFilm)
                {
                    string kWords = MyConvert.ZHLC("即涂");
                    if (MyConvert.ZHLC(ProduceRemark).Contains(kWords))
                    {
                        var v = from a in EstimateProcessPriceList
                                where (a.MachineCode == this.MachineCode || a.MachineCode.IsEmpty()) && a.ProcessCode == this.ProcessCode && MyConvert.ZHLC(a.Name).Contains(kWords)
                                orderby a.Sort_ID
                                select a;
                        if (v.IsNotNull() && v.Count() > 0)
                        {
                            EstimateProcessPriceItem fProcessItem = v.FirstOrDefault();
                            FormulaString = fProcessItem.FormulaString2.IsNotEmpty() ? fProcessItem.FormulaString2 : fProcessItem.FormulaString;
                            FormulaView = fProcessItem.FormulaView2.IsNotEmpty() ? fProcessItem.FormulaView2 : fProcessItem.FormulaView;
                            UsePersonFormula = fProcessItem.UsePersonFormula;
                            LossRatio = fProcessItem.LossRatio;
                            Unit = fProcessItem.Unit;

                            Price1 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price1) : 0;
                            Price2 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price2) : 0;
                            Price3 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price3) : 0;
                            StartPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.StartPrice) : 0;
                            StartNumber = fProcessItem.StartNumber;
                            FixPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.FixPrice) : 0;
                            LossRatio = fProcessItem.LossRatio;
                            Name = fProcessItem.Name;
                        }
                    }
                    else
                    {
                        var v = from a in EstimateProcessPriceList
                                where (a.MachineCode == this.MachineCode || a.MachineCode.IsEmpty()) && a.ProcessCode == this.ProcessCode
                                orderby a.Sort_ID
                                select a;
                        if (v.IsNotNull() && v.Count() > 0)
                        {
                            EstimateProcessPriceItem fProcessItem = v.FirstOrDefault();
                            FormulaString = fProcessItem.FormulaString2.IsNotEmpty() ? fProcessItem.FormulaString2 : fProcessItem.FormulaString;
                            FormulaView = fProcessItem.FormulaView2.IsNotEmpty() ? fProcessItem.FormulaView2 : fProcessItem.FormulaView;
                            UsePersonFormula = fProcessItem.UsePersonFormula;
                            LossRatio = fProcessItem.LossRatio;
                            Unit = fProcessItem.Unit;

                            Price1 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price1) : 0;
                            Price2 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price2) : 0;
                            Price3 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price3) : 0;
                            StartPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.StartPrice) : 0;
                            StartNumber = fProcessItem.StartNumber;
                            FixPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.FixPrice) : 0;
                            LossRatio = fProcessItem.LossRatio;
                            Name = fProcessItem.Name;
                        }
                    }
                }
                else
                {
                    var v = from a in EstimateProcessPriceList
                            where (a.MachineCode == this.MachineCode || a.MachineCode.IsEmpty()) && a.ProcessCode == this.ProcessCode
                            orderby a.Sort_ID
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        EstimateProcessPriceItem fProcessItem = v.FirstOrDefault();
                        FormulaString = fProcessItem.FormulaString2.IsNotEmpty() ? fProcessItem.FormulaString2 : fProcessItem.FormulaString;
                        FormulaView = fProcessItem.FormulaView2.IsNotEmpty() ? fProcessItem.FormulaView2 : fProcessItem.FormulaView;
                        UsePersonFormula = fProcessItem.UsePersonFormula;
                        LossRatio = fProcessItem.LossRatio;
                        Unit = fProcessItem.Unit;

                        Price1 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price1) : 0;
                        Price2 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price2) : 0;
                        Price3 = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.Price3) : 0;
                        StartPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.StartPrice) : 0;
                        StartNumber = fProcessItem.StartNumber;
                        FixPrice = mti.IsNotNull() ? mti.ToMoneyType(fProcessItem.FixPrice) : 0;
                        LossRatio = fProcessItem.LossRatio;
                        Name = fProcessItem.Name;
                    }
                }
            }

            bool isNew { get; set; }

            EstimateItem Parent { get; set; }

            C1.Win.C1FlexGrid.Row CurrentGridRow { get; set; }

            public string Sort_ID
            {
                get
                {
                    return string.Format("{0}_{1}_{2}", PartID.Length > 0 ? 100 : 200,
                                            PartID.Length > 0 ? PartID : "Default",
                                            Convert.ToString(this.id).PadLeft(10, '0'));
                }
            }
            public string PartID { get; set; }
            /// <summary>
            /// ID _id
            /// </summary>
            public int id { get; set; }

            public string ProcessCode { get; set; }

            public ProcessItem Process
            {
                get
                {
                    if (Processes.Contains(ProcessCode))
                    {
                        return Processes[ProcessCode];
                    }
                    return null;
                }
            }

            public string ProcessName
            {
                get
                {
                    if (Process.IsNotNull()) return Process.Name;
                    else return string.Empty;
                }
            }

            public bool isPrint
            {
                get
                {
                    if (Process.IsNotNull())
                    {
                        return Process.isPrint;
                    }
                    return false;
                }
            }

            public bool isDiecut
            {
                get
                {
                    if (Process.IsNotNull())
                    {
                        return Process.isDiecut;
                    }
                    return false;
                }
            }

            public bool isFilm
            {
                get
                {
                    if (Process.IsNotNull())
                    {
                        return Process.isFilm;
                    }
                    return false;
                }
            }

            public bool isNotPrint
            {
                get
                {
                    return !isPrint;
                }
            }

            public string MachineCode { get; set; }

            public MachineItem Machine
            {
                get
                {
                    if (Machines.Contains(MachineCode)) return Machines[MachineCode];
                    return null;
                }
            }

            public string MachineName
            {
                get
                {
                    if (Machine.IsNotNull()) return Machine.Name;
                    return string.Empty;
                }
            }

            public int MachineType
            {
                get
                {
                    if (Machine.IsNotNull()) return Machine.Colors > 2 ? 1 : 0;
                    return 0;
                }
            }
            /// <summary>
            /// 机器类型，从机台来
            /// </summary>
            public string MachineTypeName
            {
                get
                {
                    return MachineType == 0 ? MyConvert.ZHLC("双色机") : MyConvert.ZHLC("多色机");
                }
            }

            #region 用于公式
            /// <summary> 克重
            ///
            /// </summary>
            public double PaperGW
            {
                get
                {
                    var v = from a in Parent.Material
                            where a.PartID == PartID && a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().gw;
                    }
                    return 0;
                }
            }

            /// <summary> 原纸长
            ///
            /// </summary>
            public double PaperLength
            {
                get
                {
                    var v = from a in Parent.Material
                            where a.PartID == PartID && a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().MaterialLength;
                    }
                    return 0;
                }
            }

            /// <summary> 原纸宽
            ///
            /// </summary>
            public double PaperWidth
            {
                get
                {
                    var v = from a in Parent.Material
                            where a.PartID == PartID && a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().MaterialWidth;
                    }
                    return 0;
                }
            }
            /// <summary> 开数
            ///
            /// </summary>
            public double CutNumb
            {
                get
                {
                    var v = from a in Parent.Material
                            where a.PartID == PartID && a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().CutNumb;
                    }
                    return 0;
                }
            }
            /// <summary>成品长
            ///
            /// </summary>
            public double ProductLength
            {
                get
                {
                    return Parent.Length;
                }
            }
            /// <summary> 成品宽
            ///
            /// </summary>
            public double ProductWidth
            {
                get
                {
                    return Parent.Width;
                }
            }
            /// <summary> 成品高
            ///
            /// </summary>
            public double ProductHeight
            {
                get
                {
                    return Parent.Height;
                }
            }
            /// <summary> 成品展开长
            ///
            /// </summary>
            public double ProductUnFlodLength
            {
                get
                {
                    return Parent.UnflodLength;
                }
            }
            /// <summary> 成品展开宽
            ///
            /// </summary>
            public double ProductUnFlodWidth
            {
                get
                {
                    return Parent.UnflodWidth;
                }
            }

            /// <summary> 印纸长
            ///
            /// </summary>
            public double PrintLength
            {
                get
                {
                    var v = from a in Parent.Material
                            where a.PartID == PartID && a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().Length;
                    }
                    return 0;
                }
            }

            /// <summary> 印纸宽
            ///
            /// </summary>
            public double PrintWidth
            {
                get
                {
                    var v = from a in Parent.Material
                            where a.PartID == PartID && a.Type == EstimateMaterialItem.MaterialTypeEnum.Type_Paper
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().Width;
                    }
                    return 0;
                }
            }

            /// <summary> 标准产能
            ///
            /// </summary>
            public double StdCapacity
            {
                get
                {
                    return Machine.IsNotNull() ? Machine.StdCapacity : 0;
                }
            }
            /// <summary> 标准人数
            ///
            /// </summary>
            public double StdPersonNumber
            {
                get
                {
                    return Machine.IsNotNull() ? Machine.StdPersonNumber : 0;
                }
            }
            /// <summary> 标准调机时间
            ///
            /// </summary>
            public double StdPrepareTime
            {
                get
                {
                    return Machine.IsNotNull() ? Machine.StdPerpareTime : 0;
                }
            }
            double _StdCapcityRate = 1;
            /// <summary> 标准产能系数，BOM来的。
            ///
            /// </summary>
            public double StdCapcityRate
            {
                get
                {
                    if (_StdCapcityRate == 0) _StdCapcityRate = 1;
                    return _StdCapcityRate;
                }
                set
                {
                    _StdCapcityRate = value;
                }
            }

            double _StdLossTimeRate = 1;

            /// <summary> 标准调机系数，BOM来的。
            ///
            /// </summary>
            public double StdLossTimeRate
            {
                get
                {
                    if (_StdLossTimeRate == 0) _StdLossTimeRate = 1;
                    return _StdLossTimeRate;
                }
                set
                {
                    _StdLossTimeRate = value;
                }
            }
            /// <summary> 总贴数，从BOM来的。
            /// 默认1贴，有部件必有加工。
            /// </summary>
            public double PartNumber
            {
                get
                {
                    var v = from a in Parent.Process
                            group a by a.PartID into g
                            select g.Key;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.Count();
                    }
                    return 1;
                }
            }

            public string ProductTypeCode
            {
                get
                {
                    return Parent.TypeCode;
                }
            }

            #endregion
            public string Name { get; set; }
            /// <summary>
            /// 模數
            /// </summary>
            public double ColNumb { get; set; }

            /// <summary>
            /// 产品数
            /// </summary>
            public double ProductNumb { get; set; }
            /// <summary> 生产数（印张数)（张数，没计算公式,没加损耗之前的数字）向上取整
            ///
            /// </summary>
            public double Numb
            {
                get
                {
                    double nextProductNumb = ProductNumb;
                    if (ColNumb == 0) ColNumb = 1;
                    EstimateProcessItem epi = NextProcess;
                    if (epi.IsNotNull()) nextProductNumb = (epi.Numb + epi.LossNumb) * epi.ColNumb;
                    return Math.Ceiling(nextProductNumb / ColNumb);
                }
            }
            /// <summary> 后加工计算公式后的数量，印刷是加了损耗
            ///
            /// </summary>
            public double Numb2
            {
                get
                {
                    double xn = Numb + LossNumb;
                    if (isPrint)
                        return xn;
                    else
                    {
                        if (FormulaString.IsNotEmpty())
                        {
                            using (var eng = new V8ScriptEngine())
                            {
                                eng.AddHostObject("v", this);
                                //eng.Execute(FormulaString);
                                var Value = eng.Evaluate(FormulaString);
                                //xn = xn * Convert.ToDouble(Value);
                                //return xn;
                                return Value.ConvertTo<double>(0, true);
                            }
                        }
                        else
                        {
                            return xn;
                        }
                    }
                }
            }

            /// <summary>
            /// 固定費用（版費）
            /// </summary>
            public double FixPrice { get; set; }
            /// <summary>
            /// 四色令价/单价。
            /// </summary>
            public double Price1 { get; set; }
            /// <summary>
            /// 专色令价
            /// </summary>
            public double Price2 { get; set; }
            /// <summary>
            /// 满版令价
            /// </summary>
            public double Price3 { get; set; }
            /// <summary> 四色数
            /// </summary>
            public double ColorNumb1 { get; set; }
            /// <summary>专色数
            /// </summary>
            public double ColorNumb2 { get; set; }
            /// <summary> 满版数
            /// </summary>
            public double ColorNumb3 { get; set; }
            /// <summary>起訂價格
            /// </summary>
            public double StartPrice { get; set; }
            /// <summary>起訂量
            /// </summary>
            public double StartNumber { get; set; }
            /// <summary> 版数
            ///
            /// </summary>
            public double PlateNumb { get; set; }
            /// <summary>版式
            /// </summary>
            public string PlateType { get; set; }

            /// <summary>
            /// 公式
            /// </summary>
            public string FormulaString { get; set; }
            /// <summary>
            /// 用于显示的公式
            /// </summary>
            public string FormulaView { get; set; }

            public string Unit { get; set; }

            /// <summary>
            /// 按人力计算
            /// </summary>
            public bool UsePersonFormula { get; set; }

            public double _LossNumb = 0;
            /// <summary>
            /// 损耗量
            /// </summary>
            public double LossNumb
            {
                get
                {
                    if (_LossNumb > 0) return _LossNumb;
                    double xLossNumb = 0;
                    if (LossRatio != 0) xLossNumb = Numb * LossRatio;
                    if (isPrint)
                    {
                        xLossNumb += EstimateProcessLoss[this.ProcessCode, this.Name, this.MachineType == 1, this.Numb];
                    }
                    else
                    {
                        xLossNumb += EstimateProcessLoss[this.ProcessCode, this.Name, this.Numb];
                    }
                    return xLossNumb;
                }
                set
                {
                    _LossNumb = value;
                }
            }
            /// <summary>
            /// 损耗率
            /// </summary>
            public double LossRatio { get; set; }


            public double UnitPrice
            {
                get
                {
                    return Amount / ProductNumb;
                }
            }

            private double _iAmount = 0;

            public double Amount
            {
                get
                {
                    double xNumb = Numb <= StartNumber ? StartNumber : Numb2; //小于最低起订量，以起订量计算
                    if (isPrint)  //印刷计算
                    {
                        double _xAmount1 = (Price1 * ColorNumb1) / 1000.00, //四色对开千印价
                               _xAmount2 = (Price2 * ColorNumb2) / 1000.00, //专色对开千印价
                               _xAmount3 = (Price3 * ColorNumb3) / 1000.00, //满版对开千印价
                               _xAmountPlate = FixPrice * PlateNumb; //版费
                        _iAmount = ((_xAmount1 + _xAmount2 + _xAmount3) * xNumb) + _xAmountPlate;
                    }
                    else
                    {
                        _iAmount = (Price1 * xNumb) + FixPrice;
                    }
                    if (_iAmount < StartPrice) _iAmount = StartPrice;
                    return _iAmount;
                }
                set
                {
                    _iAmount = value;
                }
            }

            public string Remark { get; set; }


            public EstimateProcessItem NextProcess
            {
                get
                {
                    if (this.PartID.IsNotEmpty())
                    {
                        var v = from a in this.Parent.Process
                                where (a.PartID.IsEmpty() || a.PartID == a.PartID) && a.id > this.id
                                orderby a.id ascending
                                select a;
                        if (v.IsNotNull() && v.Count() > 0) return v.FirstOrDefault();
                    }
                    else
                    {
                        var v = from a in this.Parent.Process
                                where a.id > this.id
                                orderby a.id ascending
                                select a;
                        if (v.IsNotNull() && v.Count() > 0) return v.FirstOrDefault();
                    }
                    return null;
                }
            }

            public EstimateProcessItem PreviousProcess
            {
                get
                {
                    if (this.PartID.IsNotEmpty())
                    {
                        var v = from a in this.Parent.Process
                                where a.PartID == a.PartID && a.id < this.id
                                orderby a.id descending
                                select a;
                        if (v.IsNotNull() && v.Count() > 0) return v.FirstOrDefault();
                    }
                    else
                    {
                        var v = from a in this.Parent.Process
                                where a.id < this.id
                                orderby a.id descending
                                select a;
                        if (v.IsNotNull() && v.Count() > 0) return v.FirstOrDefault();
                    }
                    return null;
                }
            }


        }
        [Serializable]
        public class EstimateMaterialCollection : MyBase.MyEnumerable<EstimateMaterialItem>
        {

        }
        [Serializable]
        public class EstimateMaterialItem
        {
            public EstimateMaterialItem(EstimateItem parents, MaterialTypeEnum materialType, C1.Win.C1FlexGrid.Row GridRow)
            {
                Parent = parents;
                isNew = true;
                Type = materialType;
                CurrentGridRow = GridRow;
                CurrentGridRow.UserData = this;
                var v = from a in Parent.Material
                        select a.id;
                if (v.IsNotNull() && v.Count() > 0)
                {
                    id = v.Max() + 1;
                }
                else
                {
                    id = 1;
                }
            }

            public EstimateMaterialItem(EstimateItem parents, MyData.MyDataRow dr)
            {
                Parent = parents;
                isNew = false;
                this.id = Convert.ToInt32(dr["SortID"]);
                this.Type = (MaterialTypeEnum)Convert.ToInt32(dr["Type"]);
                this.PartID = Convert.ToString(dr["PartID"]);
                this.MaterialCode = Convert.ToString(dr["MaterialCode"]);
                this.CutNumb = Convert.ToDouble(dr["CutNumb"]);
                this.ColNumb = Convert.ToDouble(dr["ColNumb"]);
                this.UnitNumb = Convert.ToDouble(dr["UnitNumb"]);
                this.Price = Convert.ToDouble(dr["Price"]);
                this.Remark = Convert.ToString(dr["Remark"]);
            }

            protected internal EstimateMaterialItem(EstimateItem parents, MaterialTypeEnum materialType)
            {
                Parent = parents;
                isNew = true;
                Type = materialType;
                var v = from a in Parent.Material
                        select a.id;
                if (v.IsNotNull() && v.Count() > 0)
                {
                    id = v.Max() + 1;
                }
                else
                {
                    id = 1;
                }
            }

            public C1.Win.C1FlexGrid.Row CurrentGridRow { get; set; }

            public int id { get; set; }

            public bool isNew { get; set; }

            public enum MaterialTypeEnum
            {
                Type_Paper = 0,
                Type_WavePaper = 1,
                Type_Material = 2
            }

            public MaterialTypeEnum Type { get; set; }

            public EstimateItem Parent { get; set; }
            /// <summary>
            /// 部件
            /// </summary>
            public string PartID { get; set; }

            public string Sort_ID
            {
                get
                {
                    return string.Format("{0}_{1}_{2}", PartID.Length > 0 ? 100 : 200,
                                            PartID.Length > 0 ? PartID : "Default",
                                            Convert.ToString(this.id).PadLeft(10, '0'));
                }
            }

            public string MaterialCode { get; set; }

            public string MaterialName
            {
                get
                {
                    if (MaterialItem.IsNotNull())
                    {
                        return MaterialItem.Name;
                    }
                    return string.Empty;
                }
            }


            private IMaterialItem _MaterialItem;

            public IMaterialItem MaterialItem
            {
                get
                {
                    if (_MaterialItem.IsNotNull() && _MaterialItem.Code == MaterialCode)
                    {
                        return _MaterialItem;
                    }
                    else
                    {
                        if (MyServiceLogic.Materials.Contains(MaterialCode))
                        {
                            _MaterialItem = MyServiceLogic.Materials[MaterialCode];
                        }
                        else
                        {
                            _MaterialItem = MyServiceLogic.Products[MaterialCode];
                        }
                        return _MaterialItem;
                    }
                }
            }

            public string MaterialSize
            {
                get
                {
                    if (MaterialItem.IsNotNull())
                    {
                        return MaterialItem.Size;
                    }
                    return string.Empty;
                }
            }

            public double gw
            {
                get
                {
                    if (MaterialItem.IsNotNull() && MaterialItem.Type.Type == MaterialType.MaterialTypeEnum.Paper)
                    {
                        return ((MaterialItem)MaterialItem).GW;
                    }
                    return 0;
                }
            }

            public string Unit
            {
                get
                {
                    if (MaterialItem.IsNotNull())
                    {
                        if (MaterialItem.TypeCode == "403")
                        {
                            return "g";
                        }
                        else
                        {
                            return MaterialItem.Unit;
                        }
                    }
                    return string.Empty;
                }
            }

            /// <summary> 模数
            ///
            /// </summary>
            public double ColNumb { get; set; }
            /// <summary> 开数
            ///
            /// </summary>
            public double CutNumb { get; set; }

            /// <summary> 印纸长
            ///
            /// </summary>
            public double Length { get; set; }
            /// <summary>  印纸宽
            ///
            /// </summary>
            public double Width { get; set; }

            public double MaterialLength
            {
                get
                {
                    if (MaterialItem.IsNotNull() && MaterialItem.Type.Type != MaterialType.MaterialTypeEnum.Product) return ((MaterialItem)MaterialItem).Length;
                    return 0;
                }
            }

            public double MaterialWidth
            {
                get
                {
                    if (MaterialItem.IsNotNull() && MaterialItem.Type.Type != MaterialType.MaterialTypeEnum.Product) return ((MaterialItem)MaterialItem).Width;
                    return 0;
                }
            }
            /// <summary> 单位用量（用于物料）
            ///
            /// </summary>
            public double UnitNumb { get; set; }

            /// <summary> 价格，纸张是吨价。
            ///
            /// </summary>
            public double Price { get; set; }

            /// <summary>
            /// 加载默认单价
            /// </summary>
            public void LoadPrice()
            {
                Price = 0;
                MaterialPriceItem mpi = MaterialPrice[this.MaterialCode];
                MoneyTypeItem mti = null;
                double xPrice = 0;
                if (mpi.IsNotNull())
                {
                    xPrice = mpi.Price;
                    mti = MoneyType[mpi.MoneyTypeCode];
                }
                if (xPrice <= 0)
                {
                    if (this.Type == MaterialTypeEnum.Type_Paper)
                    {
                        MaterialItem thisMaterial = (MaterialItem)this.MaterialItem;
                        string xName = string.Format(@"{0}", thisMaterial.PureName.ReplaceWith(@"[\s\r\n]", string.Empty).Trim().ToUpper()).ReplaceWith(@"\(\d.+\)", string.Empty);
                        var v = from a in MaterialPrice
                                join b in Materials on a.Code equals b.Code
                                where b.GW == this.gw && b.PureName.ReplaceWith(@"[\s\r\n\(\)]", "").Trim().ToUpper().IsMatchingTo(xName) && a.Price > 450
                                orderby a.Price descending
                                select a;
                        if (v.IsNotNull() && v.Count() > 0)
                        {
                            mpi = v.FirstOrDefault();
                            xPrice = mpi.Price;
                            mti = MoneyType[mpi.MoneyTypeCode];
                        }
                    }
                }
                if (mpi.IsNotNull())
                {
                    string moneytypecode = Parent.Parent.MoneyTypeCode;
                    Price = mti.ToMoneyType(mpi.Price, moneytypecode);
                    if (this.Type == MaterialTypeEnum.Type_Paper)
                    {
                        if (mpi.MoneyTypeCode == MoneyTypeCode_Default)
                        {
                            Price = Price * 1.05;
                        }
                    }
                    return;
                }
            }

            /// <summary> 输入数量，产品数
            ///
            /// </summary>
            public double ProductNumb { get; set; }
            /// <summary>
            /// 领纸张数、物料数量
            /// </summary>
            public double Numb
            {
                get
                {
                    if (Type == MaterialTypeEnum.Type_Paper)
                    {
                        if (ColNumb == 0) ColNumb = 1;
                        return Math.Ceiling(ProductNumb / ColNumb);
                    }
                    else if (Type == MaterialTypeEnum.Type_WavePaper)
                    {
                        if (ColNumb == 0) ColNumb = 1;
                        return ProductNumb / ColNumb;
                    }
                    else
                    {
                        if (UnitNumb == 0) UnitNumb = 1;
                        return ProductNumb * UnitNumb;
                    }
                }
            }
            /// <summary>
            /// 领料张数
            /// </summary>
            public double Numb2
            {
                get
                {
                    if (Type == MaterialTypeEnum.Type_Paper)
                    {
                        if (CutNumb == 0) CutNumb = 1;
                        return Math.Ceiling((Numb + PrintLossNumb + PostprocessLossNumb) / CutNumb);
                    }
                    else if (Type == MaterialTypeEnum.Type_WavePaper)
                    {
                        return Numb;
                    }
                    else
                    {
                        if (this.MaterialItem.IsNotNull() && this.MaterialItem.TypeCode == "403")
                        {
                            return Numb / 1000.00;
                        }
                        return Numb;
                    }
                }
            }


            public double PrintLossNumb
            {
                get
                {
                    var v = from a in Parent.Process
                            where a.PartID == this.PartID && a.isPrint
                            select a;
                    if (v.IsNotNull() & v.Count() > 0)
                    {
                        return v.FirstOrDefault().LossNumb;
                    }
                    return 0;
                }
            }

            public double PostprocessLossNumb
            {
                get
                {
                    var v = from a in Parent.Process
                            where a.PartID == this.PartID && a.isNotPrint
                            select a;
                    if (v.IsNotNull() & v.Count() > 0)
                    {
                        return v.Sum(x => x.LossNumb);
                    }
                    return 0;
                }
            }

            /// <summary>
            /// 按照输入数量的金额
            /// </summary>
            public double Amount
            {
                get
                {
                    double i = 0;
                    if (MaterialItem.IsNotNull())
                    {
                        i = MaterialItem.PriceUnitConvertRatio; //采购单位转换系数。
                        return Numb2 * (Price * i);
                    }
                    else
                    {
                        return (Numb2 * Price);
                    }
                }
            }
            /// <summary>
            /// 在当前数量下的单PCS产品价格。
            /// </summary>
            public double UnitPrice
            {
                get
                {
                    return ProductNumb != 0 ? Amount / ProductNumb : 0;
                }
            }

            public string Remark { get; set; }

        }
        [Serializable]
        public class EstimateChargeCollection : MyBase.MyEnumerable<EstimateChargeItem>
        {

        }
        [Serializable]
        public class EstimateChargeItem
        {

            public EstimateChargeItem(EstimateItem parents, C1.Win.C1FlexGrid.Row GridRow)
            {
                Parent = parents;
                isNew = true;
                CurrentGridRow = GridRow;
                CurrentGridRow.UserData = this;
                var v = from a in Parent.Material
                        select a.id;
                if (v.IsNotNull() && v.Count() > 0)
                {
                    id = v.Max() + 1;
                }
                else
                {
                    id = 1;
                }
            }

            public EstimateChargeItem(EstimateItem parents, MyData.MyDataRow dr)
            {
                Parent = parents;
                isNew = false;
                this.id = Convert.ToInt32(dr["PID"]);
                this.PartID = Convert.ToString(dr["PartID"]);
                this.Name = Convert.ToString(dr["Name"]);
                this.Price = Convert.ToDouble(dr["Price"]);
                this.Unit = Convert.ToString(dr["Unit"]);
                this.FixPrice = Convert.ToDouble(dr["FixPrice"]);
                this.FormulaString = Convert.ToString(dr["FormulaString"]);
                this.FormulaView = Convert.ToString(dr["FormulaView"]);
                this.Remark = Convert.ToString(dr["Remark"]);
            }

            public C1.Win.C1FlexGrid.Row CurrentGridRow { get; set; }

            public int id { get; set; }

            public bool isNew { get; set; }

            public EstimateItem Parent { get; set; }
            /// <summary>
            /// 部件
            /// </summary>
            public string PartID { get; set; }

            public string Name { get; set; }
            /// <summary> 价格1（×单价）
            ///
            /// </summary>
            public double Price { get; set; }
            /// <summary> 系数（单位用量）
            ///
            /// </summary>
            public double UnitNumber { get; set; }
            /// <summary> 单位
            ///
            /// </summary>
            public string Unit { get; set; }
            /// <summary> 价格2（加单价）
            ///
            /// </summary>
            public double FixPrice { get; set; }
            /// <summary> 公式（显示）
            ///
            /// </summary>
            public string FormulaView { get; set; }
            /// <summary>
            /// 公式（计算）
            /// </summary>
            public string FormulaString { get; set; }

            public double ProductNumb { get; set; }

            public double Numb
            {
                get
                {
                    return ProductNumb * UnitNumber;
                }
            }

            public double UnitPrice
            {
                get
                {
                    return ProductNumb != 0 ? Amount / ProductNumb : 0;
                }
            }

            public double Amount
            {
                get
                {
                    return (Numb * Price) + FixPrice;
                }

            }


            public string Remark { get; set; }

        }

        /// <summary>
        /// 存储器
        /// </summary>
        static private List<EstimateProcessPriceItem> _EstimateProcessPriceList = new List<EstimateProcessPriceItem>();

        static public List<EstimateProcessPriceItem> EstimateProcessPriceList
        {
            get
            {
                if (_EstimateProcessPriceList.IsNull() || _EstimateProcessPriceList.Count <= 0)
                {
                    LoadEstimateProcessList();
                }
                return _EstimateProcessPriceList;
            }
        }

        static public void ClearEstimateProcessList()
        {
            _EstimateProcessPriceList = null;
        }
        /// <summary>
        /// 重新加载工序价格表
        /// </summary>
        static public void LoadEstimateProcessList()
        {
            ClearEstimateProcessList();
            string SQL = @"Select a.[_id],a.Name,a.FixPrice,a.Price1,a.Price2,a.Price3,a.StartNumber,a.StartPrice,a.Remark,a.MachineCode,a.MachineType,b.FormulaString,b.FormulaView,b.UsePersonFormula,b.LossRatio,b.Unit,b.Code as ProcessCode,a.FormulaView as FormulaView2,a.FormulaString as FormulaString2,a.Unit as Unit2,a.BeginNumb,a.Type2 from [_PMC_Process_Price] a Inner Join [moProcedure] b On a.ProcessCode = b.Code Order by [_id]";
            MyData.MyParameter[] mps = new MyData.MyParameter[]
            {
                new MyData.MyParameter("@ProcessCode", string.Empty)
            };
            using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, mps))
            {
                if (md.IsNotNull() && md.MyRows.Count > 0)
                {
                    var v = from a in md.MyRows
                            select new EstimateProcessPriceItem(a);
                    _EstimateProcessPriceList = v.ToList();
                }
            }
        }

        [Serializable]
        public class EstimateProcessPriceItem
        {
            public EstimateProcessPriceItem(MyData.MyDataRow r)
            {
                Price1 = Convert.ToDouble(r["Price1"]);
                Price2 = Convert.ToDouble(r["Price2"]);
                Price3 = Convert.ToDouble(r["Price3"]);
                this.StartPrice = Convert.ToDouble(r["StartPrice"]);
                this.StartNumber = Convert.ToDouble(r["StartNumber"]);
                this.FixPrice = Convert.ToDouble(r["FixPrice"]);
                this.Name = Convert.ToString(r["Name"]);
                this.MachineCode = Convert.ToString(r["MachineCode"]);
                this.ProcessCode = Convert.ToString(r["ProcessCode"]);
                this.MachineType = Convert.ToInt32(r["MachineType"]);

                FormulaString = Convert.ToString(r["FormulaString"]);
                FormulaView = Convert.ToString(r["FormulaView"]);

                FormulaString2 = Convert.ToString(r["FormulaString2"]);
                FormulaView2 = Convert.ToString(r["FormulaView2"]);

                UsePersonFormula = Convert.ToBoolean(r["UsePersonFormula"]);
                LossRatio = Convert.ToDouble(r["LossRatio"]);

                string Unit2 = Convert.ToString(r["Unit2"]);

                Unit = Unit2.IsNotEmpty() ? Unit2 : Convert.ToString(r["Unit"]);

                BeginNumb = Convert.ToDouble(r["BeginNumb"]);
                Type2 = Convert.ToInt32(r["Type2"]);

                id = Convert.ToInt32(r["_id"]);

            }
            public string FormulaString { get; set; }

            public string FormulaView { get; set; }

            public string FormulaString2 { get; set; }

            public string FormulaView2 { get; set; }

            public bool UsePersonFormula { get; set; }

            public double LossRatio { get; set; }

            public string Unit { get; set; }

            public double Price1 { get; set; }

            public double Price2 { get; set; }

            public double Price3 { get; set; }

            public double StartPrice { get; set; }

            public double StartNumber { get; set; }

            public double FixPrice { get; set; }

            public string Name { get; set; }

            public int MachineType { get; set; }

            public string MachineCode { get; set; }

            public string ProcessCode { get; set; }

            public int Sort_ID
            {
                get
                {
                    if (MachineCode.IsEmptyOrWhiteSpace())
                    {
                        return 900000 + id;
                    }
                    else
                    {
                        return 100000 + id;
                    }
                }
            }

            public int id { get; set; }
            /// <summary>
            /// 起始数量范围
            /// </summary>
            public double BeginNumb { get; set; }
            /// <summary>
            /// 数量范围类型,模切，1,面积，2,模数
            /// </summary>
            public int Type2 { get; set; }
        }


        #region 工序报价的损耗表
        public class EstimateProcessLossItem
        {
            public EstimateProcessLossItem(MyData.MyDataRow r)
            {
                Name = Convert.ToString(r["Name"]);
                BeginNumb = Convert.ToDouble(r["BeginNumb"]);
                ProcessCode = Convert.ToString(r["ProcessCode"]);
                Numb = Convert.ToDouble(r["Numb"]);
                MachineType = r["MachineType"].ConvertTo<bool>(false, true);
            }
            public double Price { get; set; }
            public string Name { get; set; }
            public string ProcessCode { get; set; }
            public bool MachineType { get; set; }
            /// <summary>
            /// 数量范围
            /// </summary>
            public double BeginNumb { get; set; }
            public double Numb { get; set; }

        }
        public class EstimateProcessLossCollection : MyBase.MyEnumerable<EstimateProcessLossItem>
        {
            public double this[string ProcessCode, string Name, double Numb]
            {
                get
                {
                    var v = from a in _data
                            where a.ProcessCode == ProcessCode && a.Name == Name && a.BeginNumb < Numb
                            orderby a.BeginNumb descending
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().Numb;
                    }
                    return 0;
                }
            }
            /// <summary>
            /// 用于印刷
            /// </summary>
            /// <param name="ProcessCode"></param>
            /// <param name="Name"></param>
            /// <param name="MachineType"></param>
            /// <returns></returns>
            public double this[string ProcessCode, string Name, bool MachineType, double Numb]
            {
                get
                {
                    var v = from a in _data
                            where a.ProcessCode == ProcessCode && a.Name == Name && a.MachineType == MachineType && a.BeginNumb < Numb
                            orderby a.BeginNumb descending
                            select a;
                    if (v.IsNotNull() && v.Count() > 0)
                    {
                        return v.FirstOrDefault().Numb;
                    }
                    return 0;
                }
            }
        }

        public class EstimateProcessLossList : EstimateProcessLossCollection
        {
            public EstimateProcessLossList()
            {
            }

            public void Load()
            {
                string SQL = @"
Select * from _PMC_Process_Price_Loss";
                using (MyData.MyDataTable dt = new MyData.MyDataTable(SQL))
                {
                    var v = from a in dt.MyRows
                            select new EstimateProcessLossItem(a);
                    this._data = v.ToList();
                }
            }
        }

        #endregion

        #region 工序报价放损表
        private static EstimateProcessLossList _ProcessLossForEstimate;
        /// <summary>
        /// 所有物料，纸张物料瓦楞纸
        /// </summary>
        public static EstimateProcessLossList EstimateProcessLoss
        {
            get
            {
                if (_ProcessLossForEstimate == null)
                {
                    _ProcessLossForEstimate = new EstimateProcessLossList();
                }
                return _ProcessLossForEstimate;
            }
        }
        #endregion

        #endregion

        #region 客户类

        private static Customer _Customer;

        public static Customer Customers
        {
            get
            {
                if (_Customer == null) _Customer = new Customer();
                return _Customer;
            }
        }

        /// <summary>
        /// 客户类
        /// </summary>
        [Serializable]
        public class Customer : CustomerCollection
        {
            public Customer()
            {
                reLoad();
            }

            public void reLoad()
            {
                MyRecord.Say("加载客户信息");
                string SQL = @"
Select a.*,
        b.cname as MoneyTypeName,c.name as TypeName,d.Name as DeliveryTypeName,e.Name as PaymentTypeName
from coCustomer a Left Outer Join pbMoneyType b ON a.moneytype = b.code 
                Left Outer Join coClientType c ON a.type = c.code 
                Left Outer Join coSendType d ON a.SendType=d.Code
                Left Outer Join coPayType e ON a.PayTime=e.Code
";
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL))
                {
                    reLoad(md.MyRows);
                }
            }

            static public new void Clear()
            {
                _Customer = null;
            }

        }

        /// <summary>
        /// 客户集合类
        /// </summary>
        [Serializable]
        public class CustomerCollection : MyBase.MyEnumerable<CustomerItem>
        {
            public List<CustomerSendAddressItem> AllCustomerSendAddressList { get; set; }
            public CustomerCollection()
            {
            }

            public CustomerCollection(MyData.MyDataRowColl CustomerData)
            {
                reLoad(CustomerData);
            }

            public CustomerCollection(IEnumerable<CustomerItem> Customers)
            {
                reLoad(Customers);
            }

            public void reLoad(MyData.MyDataRowColl CustomerData)
            {
                var v = from a in CustomerData
                        select new CustomerItem(a);
                reLoad(v);
            }

            public void reLoad(IEnumerable<CustomerItem> Customers)
            {
                _data = Customers.ToList();
            }

            public CustomerItem this[string CustCode]
            {
                get
                {
                    return FindByCode(CustCode);
                }
                set
                {
                    CustomerItem curItem = FindByCode(CustCode);
                    if (_current_index >= 0 && _current_index <= Length) curItem = value;
                }
            }

            public CustomerItem FindByCode(string CustCode)
            {
                CustomerItem rItem = null;
                int i = _data.FindIndex(a => a.Code.Trim().ToUpper() == CustCode.Trim().ToUpper());
                if (i >= 0)
                {
                    rItem = this[i];
                    _current_index = i;
                }
                return rItem;
            }

            public bool Contains(string CustCode)
            {
                var v = from a in _data
                        where a.Code == CustCode
                        select a.Code;
                return (v != null && v.Count() > 0);
            }
        }

        /// <summary>
        /// 客户实体类
        /// </summary>
        [Serializable]
        public class CustomerItem
        {
            public CustomerItem(MyData.MyDataRow CustomerDataRow)
            {
                if (CustomerDataRow != null)
                {
                    MyData.MyDataRow m = CustomerDataRow;
                    id = Convert.ToInt32(m["_id"]);
                    xid = Convert.ToInt32(m["id"]);
                    this.TypeCode = Convert.ToString(m["Type"]);
                    this.TypeName = Convert.ToString(m["TypeName"]);
                    Name = Convert.ToString(m["Name"]);
                    Code = Convert.ToString(m["Code"]).Trim().ToUpper();
                    Address = Convert.ToString(m["SendAddress"]);
                    if (Address == null || Address.Length <= 0) Address = Convert.ToString(m["Address"]);
                    SalesMan = Convert.ToString(m["BusinessMan"]);
                    Assistant = Convert.ToString(m["Assistant"]);
                    OEMCustomer = Convert.ToString(m["OEMCustomer"]);
                    CGUID = Guid.Parse(Convert.ToString(m["CGUID"]));
                    CreateDate = Convert.ToDateTime(m["CreateDate"]);
                    if (CreateDate <= DateTime.MinValue) CreateDate = Convert.ToDateTime(m["setDate"]);
                    Creator = Convert.ToString(m["Creator"]);
                    setDate = Convert.ToDateTime(m["setDate"]);
                    PaymentTypeCode = Convert.ToString(m["paytime"]);
                    PaymentTypeName = Convert.ToString(m["PaymentTypeName"]);
                    MoneyTypeCode = Convert.ToString(m["moneytype"]);
                    MoneyTypeName = Convert.ToString(m["MoneyTypeName"]);
                    TelNumber = Convert.ToString(m["tel"]);
                    FaxNumber = Convert.ToString(m["fax"]);
                    Connector = Convert.ToString(m["connecter"]);
                    IntelAddress = Convert.ToString(m["inteladdress"]);
                    EMailAddress = Convert.ToString(m["email"]);
                    Status = Convert.ToInt32(m["status"]);
                    Remark = Convert.ToString(m["remark"]);
                    DeliveryTypeCode = Convert.ToString(m["SendType"]);
                    DeliveryTypeName = Convert.ToString(m["DeliveryTypeName"]);
                    DepartmentCode = Convert.ToString(m["DepartmentCode"]);
                }
            }

            public void LoadSendAddressList()
            {
                string SQL = "Select * from [_SD_Customer_SendAddress] Where Code = @Code                 ";
                using (MyData.MyDataTable SendAddressDataSource = new MyData.MyDataTable(SQL, new MyData.MyParameter("@Code", this.Code)))
                {
                    var v = from a in SendAddressDataSource.MyRows
                            where Convert.ToString(a["Code"]) == this.Code
                            orderby Convert.ToInt32(a["SortID"])
                            select new CustomerSendAddressItem(a);
                    _SendAddressList = v.ToList();
                }
            }

            /// <summary>
            /// ID  手工
            /// </summary>
            public int id { get; set; }
            /// <summary>
            /// 自动ID
            /// </summary>
            public int xid { get; set; }
            /// <summary>
            /// GUID
            /// </summary>
            public Guid CGUID { get; set; }
            /// <summary>
            /// 客户编号
            /// </summary>
            public string Code { get; set; }
            /// <summary>
            /// 客户名称
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 客户分类号
            /// </summary>
            public string TypeCode { get; set; }
            /// <summary>
            /// 客户分类名称
            /// </summary>
            public string TypeName { get; set; }
            /// <summary>
            /// 供应商地址，对应SendAddress或者Address
            /// </summary>
            public string Address { get; set; }
            /// <summary>
            /// 业务员
            /// </summary>
            public string SalesMan { get; set; }
            /// <summary>
            /// 业务助理
            /// </summary>
            public string Assistant { get; set; }
            /// <summary>
            /// 付款方式编号
            /// </summary>
            public string PaymentTypeCode { get; set; }
            /// <summary>
            /// 付款方式
            /// </summary>
            public string PaymentTypeName { get; set; }
            /// <summary>
            /// 币种编号
            /// </summary>
            public string MoneyTypeCode { get; set; }
            /// <summary>
            /// 币种
            /// </summary>
            public string MoneyTypeName { get; set; }
            /// <summary>
            /// 电话
            /// </summary>
            public string TelNumber { get; set; }
            /// <summary>
            /// 传真
            /// </summary>
            public string FaxNumber { get; set; }
            /// <summary>
            /// 联系人
            /// </summary>
            public string Connector { get; set; }
            /// <summary>
            /// 手机
            /// </summary>
            public string MobileNumber { get; set; }
            /// <summary>
            /// 网址
            /// </summary>
            public string IntelAddress { get; set; }
            /// <summary>
            /// 电子邮箱
            /// </summary>
            public string EMailAddress { get; set; }
            /// <summary>
            /// 建立日期1
            /// </summary>
            public DateTime setDate { get; set; }
            /// <summary>
            ///  状态
            /// </summary>
            public int Status { get; set; }
            /// <summary>
            /// 备注
            /// </summary>
            public string Remark { get; set; }
            /// <summary>
            /// 送货方式编号
            /// </summary>
            public string DeliveryTypeCode { get; set; }
            /// <summary>
            /// 送货方式
            /// </summary>
            public string DeliveryTypeName { get; set; }

            /// <summary>
            /// OEM厂商的客户。S062。
            /// </summary>
            public string OEMCustomer { get; set; }
            /// <summary>
            /// 输入人
            /// </summary>
            public string Creator { get; set; }
            /// <summary>
            /// 输入时间
            /// </summary>
            public DateTime CreateDate { get; set; }
            /// <summary>
            /// 部门/组编号
            /// </summary>
            public string DepartmentCode { get; set; }
            /// <summary>
            /// 是否新建产品
            /// </summary>
            public bool isNew { get; private set; }

            private List<CustomerSendAddressItem> _SendAddressList = new List<CustomerSendAddressItem>();
            /// <summary>
            /// 送货地址列表
            /// </summary>
            public List<CustomerSendAddressItem> SendAddressList
            {
                get
                {
                    if (!(_SendAddressList != null && _SendAddressList.Count > 0))
                    {
                        LoadSendAddressList();
                    }
                    return _SendAddressList;
                }
            }

            public void reLoadCode()
            {
                string SQL = @"
Select Code,id,_id,setdate From coVender a Where a.CGUID=@CGUID
";
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@CGUID", CGUID, MyData.MyDataParameter.MyDataType.UniqueIdentifier)))
                {
                    if (md != null && md.MyRows.Count > 0)
                    {
                        MyData.MyDataRow mr = md.MyRows.FirstOrOnlyRow;
                        Code = Convert.ToString(mr["Code"]);
                        id = Convert.ToInt32(mr["_id"]);
                        xid = Convert.ToInt32(mr["id"]);
                        CreateDate = Convert.ToDateTime(mr["setdate"]);
                    }
                }
            }

            public bool Save()
            {
                if (CGUID == Guid.Empty) return false;
                MyData.MyCommand md = new MyData.MyCommand();
                string MainTableName = "coCustomer";
                if (isNew)
                {
                    string InsertSQL = @"
Declare @id int
Select @id=(Select IsNull(max(id),0) +1 From {0})
Insert Into {0}(id,CGUID,Code,Creator,Type) Values (@id,@CGUID,@Code,@Inputer,@Type)";
                    InsertSQL = string.Format(InsertSQL, MainTableName);
                    MyData.MyDataParameter[] mpInsert = new MyData.MyDataParameter[]
                    {
                        new MyData.MyDataParameter("@CGUID",CGUID, MyData.MyDataParameter.MyDataType.UniqueIdentifier),
                        new MyData.MyDataParameter("@Inputer",LoginUser.Name),
                        new MyData.MyDataParameter("@Type",TypeCode),
                        new MyData.MyDataParameter("@Code",Code),
                    };
                    md.Add(InsertSQL, "Insert", mpInsert);
                }
                string SQL = @" Update a Set 
                                  setdate=(Case When setdate is Null Then GetDate() Else setdate End),
                                  Creator=(Case When Creator is Null Then @Editor Else Creator End),
                                  name = @name,
                                  Code = @Code,
                                  type = @type,
                                  paytime = @paytime,
                                  moneytype = @moneytype,
                                  address = @address,
                                  SendAddress = @address,
                                  tel = @tel,
                                  fax = @fax,
                                  postcode = @postcode,
                                  connecter = @connecter,
                                  inteladdress = @inteladdress,
                                  email = @email,
                                  status = 0,
                                  remark = @remark,
                                  BankNo = Null,BankName = Null,rate2 = Null,
                                  Assistant = @Assistant,
                                  BusinessMan = @BusinessMan,
                                  OEMCustomer = @OEMCustomer,
                                  SendType = @DeliveryType,
                                  MobileNumber = @MobileNumber,
                                  DepartmentCode = @DepartmentCode,
                                  ERPTYPE = 'NEWERP'
                            From {0} a Where a.CGUID=@CGUID
";
                SQL = string.Format(SQL, MainTableName);

                MyData.MyDataParameter[] mpUpdate = new MyData.MyDataParameter[]
                {
                    new MyData.MyDataParameter("@Editor",MyServiceLogic.LoginUser.Name),
                    new MyData.MyDataParameter("@name",Name),
                    new MyData.MyDataParameter("@Code",Code),
                    new MyData.MyDataParameter("@type",TypeCode),
                    new MyData.MyDataParameter("@paytime",PaymentTypeCode),
                    new MyData.MyDataParameter("@address",Address),
                    new MyData.MyDataParameter("@tel",TelNumber),
                    new MyData.MyDataParameter("@fax",FaxNumber),
                    new MyData.MyDataParameter("@postcode",string.Empty),
                    new MyData.MyDataParameter("@connecter",Connector),
                    new MyData.MyDataParameter("@inteladdress",IntelAddress),
                    new MyData.MyDataParameter("@email",EMailAddress),
                    new MyData.MyDataParameter("@remark",Remark),
                    new MyData.MyDataParameter("@moneytype",MoneyTypeCode ),
                    new MyData.MyDataParameter("@Assistant",Assistant),
                    new MyData.MyDataParameter("@BusinessMan",SalesMan),
                    new MyData.MyDataParameter("@DeliveryType",DeliveryTypeCode),
                    new MyData.MyDataParameter("@OEMCustomer",OEMCustomer),
                    new MyData.MyDataParameter("@MobileNumber",MobileNumber),
                    new MyData.MyDataParameter("@DepartmentCode",DepartmentCode),
                    new MyData.MyDataParameter("@CGUID",CGUID, MyData.MyDataParameter.MyDataType.UniqueIdentifier)
                };
                md.Add(SQL, "Update", mpUpdate);

                string SQLDelete = @"Delete From [_SD_Customer_SendAddress] Where Code = @Code ";
                md.Add(SQLDelete, "SQLDelete", new MyData.MyDataParameter("@Code", Code));
                string SQLDetial = @"Insert Into [_SD_Customer_SendAddress] (Inputer,zbid,Code,SendAddress,SortID)
 Values(@Inputer,@zbid,@Code,@SendAddress,@SortID)";
                if (isNew)
                {
                    SQLDetial = string.Format(@"
Declare @MainID Int 
Set @MainID=IDENT_CURRENT('{0}')
Insert Into [_SD_Customer_SendAddress] (Inputer,zbid,Code,SendAddress,SortID)
 Values(@Inputer,@MainID,@Code,@SendAddress,@SortID)", MainTableName);
                }
                if (_SendAddressList != null && _SendAddressList.Count > 0)
                {
                    foreach (var SendItem in _SendAddressList)
                    {
                        MyData.MyDataParameter[] SendParams = new MyData.MyDataParameter[]
                        {
                            new MyData.MyDataParameter("@SendAddress",SendItem.SendAddress),
                            new MyData.MyDataParameter("@SortID",SendItem.SortID, MyData.MyDataParameter.MyDataType.Int),
                            new MyData.MyDataParameter("@Inputer",MyServiceLogic.LoginUser.Name),
                            new MyData.MyDataParameter("@Code",this.Code),
                            new MyData.MyDataParameter("@zbid",this.id, MyData.MyDataParameter.MyDataType.Int)
                        };
                        md.Add(SQLDetial, string.Format("SQLDetial{0}", SendItem.SortID), SendParams);
                    }
                }
                return md.Execute();
            }
        }


        [Serializable]
        public class CustomerSendAddressItem
        {
            public CustomerSendAddressItem(MyData.MyDataRow DataRow)
            {
                MyData.MyDataRow m = DataRow;
                CGUID = Guid.Parse(Convert.ToString(m["CGUID"]));
                InputDate = Convert.ToDateTime(m["InputDate"]);
                Inputer = Convert.ToString(m["Inputer"]);
                id = Convert.ToInt32(m["_ID"]);
                zbid = Convert.ToInt32(m["zbid"]);
                Code = Convert.ToString(m["Code"]);
                SendAddress = Convert.ToString(m["SendAddress"]);
                SortID = Convert.ToInt32(m["SortID"]);
                Name = Convert.ToString(m["Name"]);
            }

            public CustomerSendAddressItem(C1.Win.C1FlexGrid.Row GridRow, string Code, int MainID)
            {
                C1.Win.C1FlexGrid.Row m = GridRow;
                CGUID = Guid.NewGuid();
                InputDate = MyServiceLogic.Now;
                Inputer = MyServiceLogic.LoginUser.Name;
                zbid = MainID;
                this.Code = Code;
                SendAddress = Convert.ToString(m["SendAddress"]);
                SortID = m.Index;
                Name = string.Empty;
            }

            public Guid CGUID { get; set; }
            public string Inputer { get; set; }
            public DateTime InputDate { get; set; }
            /// <summary>
            /// 自增长
            /// </summary>
            public int id { get; set; }
            /// <summary>
            /// 客户信息表的 [_id]
            /// </summary>
            public int zbid { get; set; }

            public string Code { get; set; }

            public string SendAddress { get; set; }

            public int SortID { get; set; }

            public string Name { get; set; }
        }

        #endregion


    }
}
/*
自动计算BOM的价格。
*/