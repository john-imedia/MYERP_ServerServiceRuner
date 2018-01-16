using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;
using System.Web;
using MYERP_ServerServiceRuner.Base;
using System.IO;
using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;


namespace MYERP_ServerServiceRuner
{
    partial class MainService
    {
        #region 所有产品

        private static Product _Product = null;

        public static Product Products
        {
            get
            {
                if (_Product == null)
                {
                    _Product = new Product();
                }
                else if (_Product.needReload)
                {
                    _Product.reLoad();
                }
                return _Product;
            }
        }
        #endregion
        #region 所有物料类
        private static Material _Materials;
        /// <summary>
        /// 所有物料，纸张物料瓦楞纸
        /// </summary>
        public static Material Materials
        {
            get
            {
                if (_Materials == null)
                {
                    _Materials = new Material();
                }
                else if (_Materials.needReload)
                {
                    _Materials.reLoad();
                }
                return _Materials;
            }
        }


        #endregion
        #region 产品类
        [Serializable]
        public class Product : ProductCollection
        {
            public enum LoadProductType { WithProcessView = 0, Slience = 1 }

            private MyData.MyDataTable _dataSource;
            /// <summary>
            /// 最后一次加载的时间
            /// </summary>
            public DateTime LastLoadTime { get; private set; }

            /// <summary>
            /// 加载的次数
            /// </summary>
            public int LoadTimes { get; private set; }

            private DateTime _LastChangeTime = DateTime.Now, TestLastChangeTime = DateTime.Now;
            /// <summary>
            /// 最后一次检查时间
            /// </summary>
            public DateTime LastChangeTime
            {
                get
                {
                    double xt = (DateTime.Now - TestLastChangeTime).TotalMinutes, periodxt = 0;
                    if (MainService.CompanyType == "MY")
                    {
                        periodxt = 5;
                    }
                    else if (CompanyType == "MD")
                    {
                        periodxt = 9;
                    }
                    else
                    {
                        periodxt = 3;
                    }
                    if (xt > periodxt)
                    {
                        string SQL = "Select Max([Date]) from [_BS_Product_StatusRecorder]";
                        object v = MyData.MyCommand.ExecuteScalar(SQL);
                        DateTime oValue;
                        if (DateTime.TryParse(v.ToString(), out oValue))
                        {
                            _LastChangeTime = oValue;
                        }
                        else
                        {
                            _LastChangeTime = DateTime.MinValue;
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
                    double periodxt1 = 0, periodxt2 = 0;
                    if (CompanyType == "MT" || CompanyType == "MD")
                    {
                        periodxt1 = 60;
                        periodxt2 = 20;
                    }
                    else
                    {
                        periodxt1 = 45;
                        periodxt2 = 15;
                    }
                    double lTimeM = (DateTime.Now - LastLoadTime).TotalMinutes;
                    //double lChangeM = (LastChangeTime - LastLoadTime).Minutes;
                    return lTimeM > periodxt1;// || lChangeM > periodxt2;
                }
            }


            /// <summary>
            /// 原始数据
            /// </summary>
            public MyData.MyDataTable DataSource
            {
                get
                {
                    return _dataSource;
                }
            }

            public Product()
                : base()
            {
                LoadingFlag = string.Empty;
                loadProduct();
            }

            private string LoadingFlag { get; set; }

            private void loadProduct(LoadProductType LdType)
            {
                bool LoadRunning = false;
                lock (LoadingFlag) LoadRunning = LoadingFlag == "Running";
                if (LoadRunning) return;
                try
                {
                    if (LdType == LoadProductType.WithProcessView)
                    {
                        MyRecord.Say("-----初始化印件资料-----");
                        MyRecord.Say("初始化所有产品到本地，请稍等。（每隔40分钟会刷新一次。）");
                    }
                    string SQL = @"Select * from _BS_AllProducts Where ISNULL(@NullParams,'') = '' ";
                    lock (LoadingFlag) LoadingFlag = "Running";
                    _dataSource = new MyData.MyDataTable(SQL, 60, new MyData.MyDataParameter("@NullParams", ""));
                    MyRecord.Say( string.Format("已经读取了{0}条印件资料，正在保存到本机电脑。", _dataSource.MyRows.Count));
                    int iCount = 1;
                    _data.Clear();
                    _data = new List<ProductItem>();
                    var vxr = from a in _dataSource.MyRows
                              orderby a.Value("Code")
                              select a;
                    foreach (var a in vxr)
                    {
                        _data.Add(new ProductItem(a));
                        iCount++;
                    }
                    MyRecord.Say("-----印件资料读取完毕。-----");

                    LoadTimes += 1;
                    _LastChangeTime = LastLoadTime = TestLastChangeTime = DateTime.Now;
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                finally
                {
                    lock (LoadingFlag) LoadingFlag = string.Empty;
                }
            }

            private void loadProduct()
            {
                loadProduct(LoadProductType.WithProcessView);
            }
            /// <summary>
            /// 默认重新加载产品，会跳出提示的。
            /// </summary>
            public void reLoad()
            {
                loadProduct(LoadProductType.WithProcessView);
            }
            /// <summary>
            /// 重新加载产品。
            /// </summary>
            /// <param name="LdType">可以设置为不显示提示</param>
            public void reLoad(LoadProductType LdType)
            {
                loadProduct(LdType);
            }


            public void DoNop()
            {
                return;
            }

            static public string ProductMaterialCatlogTypeCode
            {
                get
                {
                    if (CompanyType == "ZS")
                        return "64";
                    else
                        return "1";
                }
            }

        }

        [Serializable]
        public class ProductCollection : MyBase.MyEnumerable<ProductItem>
        {
            #region 存储器

            public List<ProductItem> data
            {
                get
                {
                    return _data;
                }
            }

            #endregion 存储器

            public ProductCollection()
            {
            }

            public ProductCollection(IEnumerable<ProductItem> EnumProd)
            {
                _data = EnumProd.ToList();
            }

            private ProductItem getItem(string code)
            {
                try
                {
                    return _data.Where(a => a.Code == code).FirstOrDefault();
                }
                catch
                {
                    return null;
                }

                //var v = from a in _data
                //        where a.Code == Code
                //        select a;
                //if (v != null && v.Count() > 0)
                //    return v.FirstOrDefault();
                //else
                //    return null;
            }

            /// <summary>
            /// 产品索引，通过产品编号（明扬产品编号）
            /// </summary>
            /// <param Name="Code">明扬产品编号 S091-109-0001</param>
            /// <returns></returns>
            public ProductItem this[string code]
            {
                get
                {
                    return getItem(code);
                }
                set
                {
                    ProductItem x = getItem(code);
                    _data.Remove(x);
                    _data.Add(value);
                }
            }

            #region 内容操作

            public ProductItem Add(string prodcode)
            {
                ProductItem proditem = new ProductItem(prodcode);
                Add(proditem);
                return proditem;
            }

            public override ProductItem Add()
            {
                ProductItem proditem = new ProductItem(string.Empty);
                Add(proditem);
                return proditem;
            }

            public bool Contains(string code)
            {
                try
                {
                    if (_data.IsNotEmptySet())
                    {
                        var v = from a in _data
                                where a.Code == code
                                select a;
                        return v.IsNotEmptySet();
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                    return false;
                }
            }

            #endregion 内容操作
        }

        [Serializable]
        public class ProductItem : IMaterialItem
        {
            #region 构造器

            public ProductItem(string prodcode)
            {
                Code = prodcode;
                isNew = (prodcode.Length == 0);
            }

            public ProductItem(MyData.MyDataRow proddatarow)
            {
                LoadProductItem(proddatarow);
            }

            /// <summary>
            /// 从数据加载
            /// </summary>
            /// <param Name="a"></param>
            private void LoadProductItem(MyData.MyDataRow a)
            {
                //if (MySimplyBasic.WaitProcess.DefaultForm.IsNotNull())
                //{
                //    MySimplyBasic.WaitProcess.DefaultForm.DescriptionWord = string.Format("正在处理：{0}", a.Value("code"));
                //}

                isNew = false;
                _ID = a.IntValue("_ID");
                Code = a.Value("code");
                Name = a.Value("name");
                CustCode = a.Value("custcode");
                TypeCode = a.Value("type");
                TypeName = a.Value("typename");
                StockLocation = a.Value("location");
                FilePath = a.Value("PathName");
                //Size = Convert.ToString(a["specs2"]);
                Length = a.Value<float>("length");
                Width = a.Value<float>("width");
                Height = a.Value<float>("height");
                UnflodLength = a.Value<float>("pLength");
                UnflodWidth = a.Value<float>("pWidth");
                PageNumberCover = a.Value<int>("CoverPage");
                PageNumberInside = a.Value<int>("ContentPage");
                BindingStyle = a.Value("PackType");
                CustProdName = a.Value("mCustNO");
                ColorSTD = a.Value("color");
                OriginalManuscript = a.Value("Manuscript");
                PlateCode = a.Value("pn");
                Prepress = a.Value("BackMark");
                ProductionMemo = a.Value("ProdRemark");
                DateTime xInputDate = a.DateTimeValue("inputdate");
                InputDate = CreateDate = xInputDate.IsNotEmpty() ? xInputDate : a.DateTimeValue("CreateDate");
                Inputer = Creator = a.Value("inputer");
                PhoneSubject = a.Value("PhoneSubject");
                ModifyDate = a.DateTimeValue("editdate");
                Modifyer = a.Value("editor");
                Description = a.Value("Description");
                mType = a.Value("mType");
                DeclaretionName = a.Value("DeclaretionName");
                Unit = a.Value("unit");
                Price = a.Value<double>("price");
                Price1 = a.Value<double>("price1");
                Price2 = a.Value<double>("price2");
                Price3 = a.Value<double>("price3");
                Price4 = a.Value<double>("price4");
                Bonded = a.BooleanValue("bonded");
                Secret = a.BooleanValue("secret");
                FSC = a.BooleanValue("FSC");
                Weight = a.Value<double>("weight") * 1000;
                MaxStock = a.IntValue("maxstock");
                MinStock = a.IntValue("minstock");
                Status = a.IntValue("status");
                Remark1 = a.Value("remark");
                Remark2 = a.Value("remark2");
                FinanceType = a.Value("financeType");
                CutPlateCode = a.Value("CutPlateCode");
                ScreenPlateCode = a.Value("ScreenPlateCode");
                HotPlateCode = a.Value("HotPlateCode");
                BomID = a.IntValue("BomID");
                BomRemark = a.Value("BomRemark");
                BomCreateDate = a.DateTimeValue("BomInputDate");
                BomCreator = a.Value("BomInputer");
                BomCheckDate = a.DateTimeValue("BomCheckDate");
                BomChecker = a.Value("BomChecker");
                MutiplePart = a.BooleanValue("mPart"); // Convert.ToInt32(a["mPart"]) == 1;
                BomStatus = a.IntValue("BomStatus");
                string xFlagCode = a.Value("FlagCode");
                isNewERPBom = xFlagCode.IsNotEmptyOrWhiteSpace(); // (Convert.ToString(a["FlagCode"]) != null && Convert.ToString(a["FlagCode"]).Length > 0);
                SampleDate = a.DateTimeValue("Smpdate");
                //SampleType = Convert.ToString(a["smp"]);
                SampleType = isSample ? MyConvert.ZHLC("无样") : MyConvert.ZHLC("内部留样");
                MassProductionPMName = a.Value("MPMan");
                MassProductionApproveDate = a.DateTimeValue("MPDate");
                MassProductionRemark = a.Value("MPRemark");
                AcceptOnDeviation = a.BooleanValue("AOD");
                InventoryTurnOverDays = a.IntValue("InventoryTurnOverDays");
                LastOutStockDate = a.DateTimeValue("LastOutStockDate");
                LotInStockDate = a.DateTimeValue("LotInStockDate");

                InStockType = a.Value("InStockType");
                PickingType = a.IntValue("PackingType");

                string x_PriceUnit = a.Value("PriceUnit");
                if (x_PriceUnit.IsEmptyOrWhiteSpace())
                {
                    PriceUnit = Unit;
                    PriceUnitConvertRatio = 1;
                }
                else
                {
                    PriceUnit = x_PriceUnit;
                    PriceUnitConvertRatio = a.Value<double>("PriceUnitConvertRatio");
                }
                string x_PurchaseUnit = a.Value("PurchaseUnit");
                if (x_PurchaseUnit.IsEmptyOrWhiteSpace())
                {
                    PurchaseUnit = Unit;
                    PurchaseUnitConvertRatio = 1;
                }
                else
                {
                    PurchaseUnit = x_PurchaseUnit;
                    PurchaseUnitConvertRatio = a.Value<double>("PurchaseUnitConvertRatio");
                }

                //specs	 custprodno		kCode	EstRdsNo	flagCode	powerrate		fPaperRemark	fColorRemark	fProceRemark	fMarkRemark
                //lPaperRemark	lColorRemark	lProceRemark	lMarkRemark	lProductRemark	CreateDate	fNumb		Smp	SmpDate
                //FM1	FM2	FM3	FM4	FM5	FM6	FM7	FM8	FM9	FM10	FM11	FM12	NW1	NW2	NW3	NW4	NW5	NW6	NW7	NW8	NW9	NW10	NW11	NW12	NW	FM
            }

            #endregion 构造器

            /// <summary>
            /// ID
            /// </summary>
            public int _ID { get; set; }

            /// <summary>
            /// 明扬产品编号
            /// </summary>
            public string Code { get; private set; }

            public bool isNew { get; set; }

            #region 自身属性

            /// <summary>
            /// 料号
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// 客户料号
            /// </summary>
            public string CustProdName { get; set; }

            /// <summary>
            /// 客户编号
            /// </summary>
            public string CustCode { get; set; }

            /// <summary>
            /// 描述，业务品名，标签品名
            /// </summary>
            public string Description { get; set; }

            /// <summary>
            /// 制造分类
            /// </summary>
            public string mType { get; set; }

            /// <summary>
            /// 报关品名,商品名
            /// </summary>
            public string DeclaretionName { get; set; }

            /// <summary>
            /// 印件种类编号
            /// </summary>
            public string TypeCode { get; set; }

            /// <summary>
            /// 印件种类名称
            /// </summary>
            public string TypeName { get; set; }

            /// <summary>
            /// 类型
            /// </summary>
            public MaterialTypeItem Type
            {
                get
                {
                    return MaterialTypes[TypeCode];
                }
            }

            /// <summary>
            /// 版号
            /// </summary>
            public string PlateCode { get; set; }
            /// <summary>
            /// 刀模编号
            /// </summary>
            public string CutPlateCode { get; set; }
            /// <summary>
            /// 烫金版号
            /// </summary>
            public string HotPlateCode { get; set; }
            /// <summary>
            /// 网版号
            /// </summary>
            public string ScreenPlateCode { get; set; }

            /// <summary>
            /// 机种
            /// </summary>
            public string PhoneSubject { get; set; }

            /// <summary>
            /// 建立日期
            /// </summary>
            public DateTime CreateDate { get; set; }

            /// <summary>
            /// 最后修改日期
            /// </summary>
            public DateTime ModifyDate { get; set; }

            /// <summary>
            /// 建立人
            /// </summary>
            public string Creator { get; set; }

            /// <summary>
            /// 最后一次修改人
            /// </summary>
            public string Modifyer { get; set; }

            /// <summary>
            /// 尺寸规格
            /// </summary>
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
                    string uSize = value;
                    Length = Width = Height = 0;
                    Regex rgx = new Regex(@"^([\d\.]+)(?=[xX*])");
                    string SL = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    rgx = new Regex(@"(?<=[xX*])([\d\.]+)(?=[xX*])");
                    string SW = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    rgx = new Regex(@"(?<=[xX*])([\d\.]+)$");
                    string SH = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    float l = 0, w = 0, h = 0;
                    if (float.TryParse(SL, out l)) Length = l;
                    if (float.TryParse(SW, out w)) Width = w;
                    if (float.TryParse(SH, out h))
                    {
                        if (w == 0) Width = h;
                        else Height = h;
                    }
                }
            }

            /// <summary>
            /// 展开尺寸规格
            /// </summary>
            public string UnfoldSize
            {
                get
                {
                    return string.Format("{0:0.##}x{1:0.##}", UnflodLength, UnflodWidth);
                }
                set
                {
                    string uSize = value;
                    UnflodLength = UnflodWidth = 0;
                    Regex rgx = new Regex(@"^([\d\.]+)(?=[xX*])");
                    string uFL = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    rgx = new Regex(@"(?<=[xX*])([\d\.]+)$");
                    string uFW = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    float ufl = 0, ufw = 0;
                    if (float.TryParse(uFL, out ufl)) UnflodLength = ufl;
                    if (float.TryParse(uFW, out ufw)) UnflodWidth = ufw;
                }
            }

            /// <summary>
            /// 成品长
            /// </summary>
            public float Length { get; set; }

            /// <summary>
            /// 成品宽
            /// </summary>
            public float Width { get; set; }

            /// <summary>
            /// 成品高
            /// </summary>
            public float Height { get; set; }

            /// <summary>
            /// 展开长
            /// </summary>
            public float UnflodLength { get; set; }

            /// <summary>
            /// 展开宽
            /// </summary>
            public float UnflodWidth { get; set; }

            /// <summary>
            /// 总页数
            /// </summary>
            public int PageNumberTotal
            {
                get
                {
                    return PageNumberCover + PageNumberInside;
                }
            }

            /// <summary>
            /// 封面页数
            /// </summary>
            public int PageNumberCover { get; set; }

            /// <summary>
            /// 内文页数
            /// </summary>
            public int PageNumberInside { get; set; }

            /// <summary>
            /// 装订方式
            /// </summary>
            public string BindingStyle { get; set; }

            /// <summary>
            /// 原稿内容
            /// </summary>
            public string OriginalManuscript { get; set; }

            /// <summary>
            /// 颜色标准
            /// </summary>
            public string ColorSTD { get; set; }

            /// <summary>
            /// 前置作业
            /// </summary>
            public string Prepress { get; set; }

            /// <summary>
            /// 特殊品质/加工说明
            /// </summary>
            public string ProductionMemo { get; set; }

            /// <summary>
            /// 产品信息设定价格
            /// </summary>
            public double Price { get; set; }
            public double Price1 { get; set; }
            public double Price2 { get; set; }
            public double Price3 { get; set; }
            public double Price4 { get; set; }

            /// <summary>
            /// 数量单位
            /// </summary>
            public string Unit { get; set; }

            /// <summary>
            /// 保税
            /// </summary>
            public bool Bonded { get; set; }

            /// <summary>
            /// 保密
            /// </summary>
            public bool Secret { get; set; }

            /// <summary>
            /// 单重
            /// </summary>
            public double Weight { get; set; }

            /// <summary>
            /// 最大库存
            /// </summary>
            public double MaxStock { get; set; }

            /// <summary>
            /// 最小库存
            /// </summary>
            public double MinStock { get; set; }

            /// <summary>
            /// 状态数字
            /// </summary>
            public int Status { get; set; }

            /// <summary>
            /// 备注说明1
            /// </summary>
            public string Remark1 { get; set; }

            /// <summary>
            /// 备注说明2
            /// </summary>
            public string Remark2 { get; set; }

            /// <summary>
            /// 产品类型——财务分类
            /// KIT成品、半成品这种。
            /// </summary>
            public string FinanceType { get; set; }

            /// <summary>
            /// 档案位置
            /// </summary>
            public string FilePath { get; set; }

            /// <summary>
            /// 前置作业完成时间
            /// </summary>
            public DateTime PrepressFinishDate { get; set; }

            /// <summary>
            /// 入仓别
            /// </summary>
            public string StockLocation { get; set; }

            /// <summary>
            /// 落版完成时间
            /// </summary>
            public DateTime PlateFinishDate { get; set; }
            /// <summary>
            /// 输入日期，和CreateDate一样
            /// </summary>
            public DateTime InputDate { get; set; }
            /// <summary>
            /// 输入人和Creator一样
            /// </summary>
            public string Inputer { get; set; }
            /// <summary>
            /// 审核日期
            /// </summary>
            public DateTime CheckDate { get; set; }
            /// <summary>
            /// 审核人
            /// </summary>
            public string Checker { get; set; }
            /// <summary>
            /// FSC
            /// </summary>
            public bool FSC { get; set; }

            /// <summary>
            /// 是否属于KIT
            /// </summary>
            public bool isKIT
            {
                get
                {
                    if (CompanyType == "ZS")
                        return false; //纸塑直接返回false
                    else if (CompanyType == "MT")
                        return TypeCode == "113";
                    else if (CompanyType == "MD")
                        return TypeCode == "113" || TypeCode == "116" || TypeCode == "120";
                    else
                        return TypeCode == "110";
                }
            }

            public bool isMY
            {
                get
                {
                    if (CompanyType == "MD")
                    {
                        return CustCode == "MY" || CustCode == "YX012";
                    }
                    else
                    {
                        return CustCode == "MY";
                    }
                }
            }

            /// <summary>
            /// 这是书刊
            /// </summary>
            public bool isBook
            {
                get
                {
                    if (CompanyType != "ZS")
                    {
                        string mBookType = string.Empty;
                        if (CompanyType == "MY")
                            mBookType = ",103,104,113,";
                        else if (CompanyType == "MT")
                            mBookType = ",103,104,107,";
                        else if (CompanyType == "MD")
                            mBookType = ",103,104,";
                        return mBookType.IndexOf(Convert.ToString(TypeCode).Trim()) > 0;
                    }
                    else
                        return false;///纸塑没有书刊
                }
            }

            /// <summary>
            /// 是否使用新ERP修改过BOM
            /// </summary>
            public bool isNewERPBom { get; set; }
            /// <summary>
            /// 样品类型
            /// </summary>
            public string SampleType { get; set; }
            /// <summary>
            /// 样品日期
            /// </summary>
            public DateTime SampleDate { get; set; }
            /// <summary>
            /// 是否有氧
            /// </summary>
            public bool isSample
            {
                get
                {
                    if (CompanyType == "MD")
                        return (SampleDate != null && SampleDate >= DateTime.Parse("1900-01-01"));
                    else
                        return (SampleDate != null && SampleDate > DateTime.Parse("2000-01-01"));
                }
            }
            /// <summary>
            /// 量产确认日期
            /// </summary>
            public DateTime MassProductionApproveDate { get; set; }
            /// <summary>
            /// 量产说明
            /// </summary>
            public string MassProductionRemark { get; set; }
            /// <summary>
            /// 量产PM确认人
            /// </summary>
            public string MassProductionPMName { get; set; }
            /// <summary>
            /// 是否可以量产
            /// </summary>
            public bool isMassProduction
            {
                get
                {
                    return (MassProductionApproveDate != null && MassProductionApproveDate > DateTime.Parse("2000-01-01"));
                }
            }
            /// <summary>
            /// 是否特批开工单。
            /// </summary>
            public bool AcceptOnDeviation { get; set; }
            /// <summary>
            /// 单价单位
            /// </summary>
            public string PriceUnit { get; set; }
            /// <summary>
            /// 单价单位换算比
            /// </summary>
            public double PriceUnitConvertRatio { get; set; }
            /// <summary>
            /// 采购单位
            /// </summary>
            public string PurchaseUnit { get; set; }
            /// <summary>
            /// 采购单位换算比
            /// </summary>
            public double PurchaseUnitConvertRatio { get; set; }
            /// <summary>
            /// 采购周期
            /// </summary>
            public double LeadTime { get; set; }
            /// <summary>
            /// 最低起订量
            /// </summary>
            public double MOQ { get; set; }
            /// <summary>
            /// 保质期,天数
            /// </summary>
            public int EXPeriod { get; set; }
            /// <summary>
            /// 呆料标志
            /// </summary>
            public bool Idel { get; set; }
            /// <summary>
            /// 呆料期限,天数
            /// </summary>
            public double IdelPeriod { get; set; }
            /// <summary>
            /// 不計庫存
            /// </summary>
            public bool isNoStock { get; set; }
            /// <summary>
            /// 這是費用
            /// </summary>
            public bool isExpenses { get; set; }
            /// <summary>
            /// 领料类型
            /// </summary>
            public int PickingType { get; set; }
            /// <summary>
            /// 入库类型
            /// </summary>
            public string InStockType { get; set; }

            /// <summary>
            /// 是否取整
            /// </summary>
            public bool NumberInteger
            {
                get
                {
                    if (CompanyType == "ZS")
                    {
                        if (TypeCode == "6102") return false;
                    }
                    return true;
                }
                set
                {

                }
            }
            /// <summary>
            /// 有效期 天
            /// </summary>
            public double Experiod { get; set; }

            /// <summary>
            /// 平均货存周转天数
            /// </summary>
            public int InventoryTurnOverDays { get; set; }
            /// <summary>
            /// 最后出库时间
            /// </summary>
            public DateTime LastOutStockDate { get; set; }

            /// <summary>
            /// 当前明细库存入库最早时间
            /// </summary>
            public DateTime LotInStockDate { get; set; }


            #endregion 自身属性

            #region Bom 属性

            /// <summary>
            /// BOM ID
            /// </summary>
            public int BomID { get; set; }
            /// <summary>
            /// BOM备注
            /// </summary>
            public string BomRemark { get; set; }
            /// <summary>
            /// BOM建立日期
            /// </summary>
            public DateTime BomCreateDate { get; set; }
            /// <summary>
            /// BOM建立人
            /// </summary>
            public string BomCreator { get; set; }
            /// <summary>
            /// 多部件
            /// </summary>
            public bool MutiplePart { get; set; }
            /// <summary>
            /// BOM状态
            /// </summary>
            public int BomStatus { get; set; }

            public DateTime BomCheckDate { get; set; }
            public string BomChecker { get; set; }

            #region 加载BOM
            private Bom _BOM = null;

            public Bom BOM
            {
                get
                {
                    LoadBom();
                    return _BOM;
                }
                set
                {
                    _BOM = value;
                }
            }

            public void LoadBom()
            {
                if (_BOM == null && Code.Length > 0)
                {
                    _BOM = new Bom(this);
                }
            }

            public void reLoad()
            {
                //                string SQL = @"Select a.*,b.id as BomID,b.DM as CutPlateCode,b.ScreenPlate as ScreenPlateCode,b.HotPlateCode,b.Remark as BomRemark,
                //                                              b.Inputer as BomInputer,b.InputDate as BomInputDate,b.mPart,b.status as BomStatus,b.FlagCode as FlagCode,
                //                                              b.Checker as BomChecker,b.CheckDate as BomCheckDate,b.MPRemark,b.MPDate,b.MPMan
                //                                       from pbProduct a Left Outer Join pbBom b On a.Code=b.Code Where isNull(b.VesionNo,1)=1 And a.Code=@Code ";
                string SQL = @"Select * from _BS_AllProducts Where Code = @Code";
                MyData.MyDataTable _dataSource = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@Code", this.Code));
                if (_dataSource.MyRows.Count > 0) LoadProductItem(_dataSource.MyRows.FirstOrOnlyRow);
                _BOM = new Bom(this.Code);
            }
            #endregion

            #endregion

            #region 明细属性


            private List<ProductMemo> _ProductMemoData = null;

            public List<ProductMemo> PartMemoList
            {
                get
                {
                    if (_ProductMemoData == null) LoadMemo();
                    return _ProductMemoData;
                }
                set
                {
                    _ProductMemoData = value;
                }
            }

            public void LoadMemo()
            {
                _ProductMemoData = new List<ProductMemo>();
                var vMemo = from a in ProductMemoList
                            where a.ZBID == this._ID
                            orderby a.PartID
                            select a;
                _ProductMemoData = vMemo.ToList();
            }

            #endregion

            /// <summary>
            /// 深拷贝
            /// </summary>
            /// <returns></returns>
            public ProductItem DeepClone()
            {
                using (Stream objectStream = new MemoryStream())
                {
                    IFormatter formatter = new BinaryFormatter();
                    formatter.Serialize(objectStream, (ProductItem)this);
                    objectStream.Seek(0, SeekOrigin.Begin);
                    return formatter.Deserialize(objectStream) as ProductItem;
                }
            }
        }
        /// <summary>
        /// 产品明细属性类，封面内文分开。
        /// </summary>
        [Serializable]
        public class ProductMemo
        {
            public int ID { get; private set; }
            public int ZBID { get; private set; }
            public string PartID { get; private set; }

            public ProductMemo(int MainID, string PartID)
            {
                ZBID = MainID;
                this.PartID = PartID;
            }
            public string PaperRemark { get; set; }
            public string PrintRemark { get; set; }
            public string PostpressRemark { get; set; }
            public string PlateRemark { get; set; }
            public int PageNumber { get; set; }
            public double CutLength { get; set; }
            public double CutWidth { get; set; }
            public string CutSize
            {
                get
                {
                    return string.Format("{0:0.#}x{1:0.#}", CutLength, CutWidth);
                }
                set
                {
                    string uSize = value;
                    Regex rgx = new Regex(@"^([\d\.]+)(?=x)");
                    string uFL = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    rgx = new Regex(@"(?<=x)([\d\.]+)$");
                    string uFW = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    float ufl = 0, ufw = 0;
                    if (float.TryParse(uFL, out ufl)) CutLength = ufl;
                    if (float.TryParse(uFW, out ufw)) CutWidth = ufw;
                }
            }
            public double WavePaperLength { get; set; }
            public double WavePaperWidth { get; set; }
            public string WavePaperSize
            {
                get
                {
                    return string.Format("{0:0.#}x{1:0.#}", WavePaperLength, WavePaperWidth);
                }
                set
                {
                    string uSize = value;
                    Regex rgx = new Regex(@"^([\d\.]+)(?=x)");
                    string uFL = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    rgx = new Regex(@"(?<=x)([\d\.]+)$");
                    string uFW = rgx.IsMatch(uSize) ? rgx.Match(uSize).Value : "";
                    float ufl = 0, ufw = 0;
                    if (float.TryParse(uFL, out ufl)) WavePaperLength = ufl;
                    if (float.TryParse(uFW, out ufw)) WavePaperWidth = ufw;
                }
            }
            public string PlateMemo { get; set; }
            public int ColNumb { get; set; }
        }

        private static List<ProductMemo> _ProductMemoList = new List<ProductMemo>();

        public static List<ProductMemo> ProductMemoList
        {
            get
            {
                if (_ProductMemoList == null) _ProductMemoList = new List<ProductMemo>();
                if (_ProductMemoList.Count <= 0)
                {
                    reLoadProductMemoListData();
                }
                return _ProductMemoList;
            }
        }

        public static void reLoadProductMemoListData()
        {
            bool wpc = false;
            MyRecord.Say("请稍等，获取印件资料。");
            MyRecord.Say("正在加载所有印件的明细资料。\r\n(只会加载这一次)");
            try
            {
                string SQL = @"Select * from _BS_Product_List where zbid>0";
                using (MyData.MyDataTable mdt = new MyData.MyDataTable(SQL))
                {
                    var v = from a in mdt.MyRows
                            select new ProductMemo(a.IntValue("zbid"), a.Value("partid"))
                            {
                                ColNumb = a.IntValue("ColNumb"),
                                CutLength = a.Value<double>("CutLength"),
                                CutWidth = a.Value<double>("CutWidth"),
                                WavePaperLength = a.Value<double>("WavePaperLength"),
                                WavePaperWidth = a.Value<double>("WavePaperWidth"),
                                PageNumber = a.IntValue("PageNumber"),
                                PaperRemark = a.Value("PaperRemark"),
                                PrintRemark = a.Value("PrintRemark"),
                                PostpressRemark = a.Value("PostpressRemark"),
                                PlateMemo = a.Value("PlateMemo"),
                                PlateRemark = a.Value("PlateRemark")
                            };
                    if (_ProductMemoList == null) _ProductMemoList = new List<ProductMemo>();
                    _ProductMemoList.Clear();
                    _ProductMemoList.AddRange(v);
                }
            }
            finally
            {
            }

        }

        /// <summary>
        /// 产品物料类
        /// </summary>
        [Serializable]
        public class ProductMaterial
        {
            public int ID { get; set; }
            public int ZBID { get; private set; }

            public string Code { get; set; }
            public string Name { get; set; }
            public float Number { get; set; }
            public string Remark { get; set; }

            public ProductMaterial(int zbid)
            {
                ZBID = zbid;
            }
        }

        #endregion 产品类
        #region 物料接口
        public interface IMaterialItem
        {
            int _ID { get; set; }
            string Code { get; }
            string Name { get; set; }
            string Size { get; set; }
            string TypeCode { get; set; }
            DateTime InputDate { get; set; }
            string Inputer { get; set; }
            MaterialTypeItem Type { get; }
            string Unit { get; set; }
            string PriceUnit { get; set; }
            double PriceUnitConvertRatio { get; set; }
            string PurchaseUnit { get; set; }
            double PurchaseUnitConvertRatio { get; set; }
            double LeadTime { get; set; }
            double Experiod { get; set; }
            double IdelPeriod { get; set; }
            double MOQ { get; set; }
            bool isNoStock { get; set; }
            /// <summary>
            /// 这是费用，无需入库
            /// </summary>
            bool isExpenses { get; set; }
            /// <summary>
            /// 0,按照排程领料，1,按照工单领料，2,其他领料。
            /// </summary>
            int PickingType { get; set; }

            string InStockType { get; set; }
            bool NumberInteger { get; set; }
            double MaxStock { get; set; }

            double MinStock { get; set; }
            /// <summary>
            /// 保稅
            /// </summary>
            bool Bonded { get; set; }
            /// <summary>
            /// FSC
            /// </summary>
            bool FSC { get; set; }
            /// <summary>
            /// 平均货存周转天数
            /// </summary>
            int InventoryTurnOverDays { get; set; }
            /// <summary>
            /// 最后出库时间
            /// </summary>
            DateTime LastOutStockDate { get; set; }
            /// <summary>
            /// 有库存的最早入库日期
            /// </summary>
            DateTime LotInStockDate { get; set; }
            /// <summary>
            /// 填寫的單價
            /// </summary>
            double Price { get; set; }
            /// <summary>
            /// 採購或者銷售的月平均单价（每月移动平均）
            /// </summary>
            double Price1 { get; set; }
            /// <summary>
            /// 产品估价/【物料最新报价（都是报价单位单价）】
            /// </summary>
            double Price2 { get; set; }
            /// <summary>
            /// 没用
            /// </summary>
            double Price3 { get; set; }
            /// <summary>
            /// 没用
            /// </summary>
            double Price4 { get; set; }
        }
        #endregion
        #region 物料分类类
        /// <summary>
        /// 物料分类类
        /// </summary>
        [Serializable]
        public class MaterialType : MaterialTypeCollection
        {
            public enum MaterialTypeEnum
            {
                /// <summary>
                /// 无分类
                /// </summary>
                None = 0,
                /// <summary>
                /// 彩印成品
                /// </summary>
                Product = 1,
                /// <summary>
                /// 彩印纸张
                /// </summary>
                Paper = 2,
                /// <summary>
                /// 彩印瓦楞纸
                /// </summary>
                WavePaper = 3,
                /// <summary>
                /// 彩印物料
                /// </summary>
                Material = 4,
                /// <summary>
                /// 蘇州名揚彩印卷筒纸
                /// </summary>
                RollPaper = 5,
                /// <summary>
                /// 纸塑成品
                /// </summary>
                PaperPlasticProduct = 61,
                /// <summary>
                /// 纸塑原材料（浆料）
                /// </summary>
                PaperPlasticMaterial_1 = 62,
                /// <summary>
                /// 纸塑辅料(添加剂)
                /// </summary>
                PaperPlasticMaterial_2 = 63,
                /// <summary>
                /// 纸槊的模具
                /// </summary>
                PaperPlasticDieMould = 64,
                /// <summary>
                /// 纸塑的外购料
                /// </summary>
                PaperPlasticOutersource = 65,
                /// <summary>
                /// 纸槊的设备
                /// </summary>
                PaperPlasticEquipment = 66,
                /// <summary>
                /// 纸塑的其他辅料
                /// </summary>
                PaperPlasticMaterial_3 = 67
            }

            public MaterialType()
            {
                reLoad();
            }

            public void reLoad()
            {
                string SQL = @"Select * from pbMaterialTypelst";
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL))
                {
                    reLoad(md.MyRows);
                }
            }
        }

        /// <summary>
        /// 产品分类集合
        /// </summary>
        [Serializable]
        public class MaterialTypeCollection : MyBase.MyEnumerable<MaterialTypeItem>
        {
            public MaterialTypeCollection()
            {
            }

            public MaterialTypeCollection(MyData.MyDataRowCollection MaterialTypeData)
            {
                reLoad(MaterialTypeData);
            }

            public MaterialTypeCollection(IEnumerable<MaterialTypeItem> MaterialTypes)
            {
                reLoad(MaterialTypes);
            }

            public void reLoad(MyData.MyDataRowCollection Data)
            {
                var v = from a in Data
                        select new MaterialTypeItem(a);
                reLoad(v);
            }

            public void reLoad(IEnumerable<MaterialTypeItem> Data)
            {
                _data = Data.ToList();
            }

            public MaterialTypeItem this[string TypeCode]
            {
                get
                {
                    return FindByCode(TypeCode);
                }
                set
                {
                    MaterialTypeItem curItem = FindByCode(TypeCode);
                    if (_current_index >= 0 && _current_index <= Length) curItem = value;
                }
            }

            public MaterialTypeItem FindByCode(string Code)
            {
                MaterialTypeItem rItem = null;
                int i = _data.FindIndex(a => a.Code == Code);
                if (i >= 0)
                {
                    rItem = this[i];
                    _current_index = i;
                }
                return rItem;
            }
        }

        /// <summary>
        /// 产品分类实体
        /// </summary>
        [Serializable]
        public class MaterialTypeItem
        {
            public MaterialTypeItem()
            {
            }

            public MaterialTypeItem(MyData.MyDataRow MaterialTypeDataRow)
            {
                if (MaterialTypeDataRow != null)
                {
                    MyData.MyDataRow m = MaterialTypeDataRow;
                    Name = Convert.ToString(m["Name"]);
                    Code = Convert.ToString(m["Code"]);
                    id = Convert.ToInt32(m["ID"]);
                    Remark = Convert.ToString(m["Remark"]);
                    TypeName = Convert.ToString(m["Type_Name"]);
                    TypeCode = Convert.ToString(m["Type"]);
                    StatusID = Convert.ToInt32(m["Status"]);

                }
            }
            /// <summary>
            /// 备注
            /// </summary>
            public string Remark { get; set; }
            /// <summary>
            /// 分类的分类名称
            /// </summary>
            public string TypeName { get; set; }
            /// <summary>
            /// 分类的分类编号
            /// </summary>
            public string TypeCode { get; set; }
            /// <summary>
            /// 状态，2以上为核销
            /// </summary>
            public int StatusID { get; set; }
            /// <summary>
            /// 分类的分类
            /// </summary>
            public MaterialType.MaterialTypeEnum Type
            {
                get
                {
                    int x = 0;
                    if (int.TryParse(TypeCode, out x))
                        return (MaterialType.MaterialTypeEnum)x;
                    else
                        return MaterialType.MaterialTypeEnum.None;
                }
            }
            /// <summary>
            /// 分类名称
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 分类ID
            /// </summary>
            public int id { get; set; }
            /// <summary>
            /// 分类编号
            /// </summary>
            public string Code { get; set; }


            public bool isHardcoverBox
            {
                get
                {
                    if (CompanyType == "ZS")
                    {
                        return Code == "118";
                    }
                    else if (CompanyType == "MT")
                    {
                        return Code == "115";
                    }
                    else if (CompanyType == "MD")
                    {
                        return Code == "120";
                    }
                    else
                        return false;
                }
            }

            public bool isKit
            {
                get
                {
                    if (CompanyType == "MY")
                    {
                        return Code == "110";
                    }
                    else if (CompanyType == "MT")
                    {
                        return Code == "113";
                    }
                    else if (CompanyType == "MD")
                    {
                        return Code == "113" || Code == "120" || Code == "116";
                    }
                    else
                        return false;
                }
            }

            public bool isNotKit
            {
                get
                {
                    return !isKit;
                }
            }
        }

        #endregion 物料分类类
        #region 所有物料类
        /// <summary>
        /// 物料类
        /// </summary>
        [Serializable]
        public class Material : MaterialCollection
        {
            public Material()
            {
                reLoad();
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
                    if ((DateTime.Now - TestLastChangeTime).TotalMinutes > 3)
                    {
                        string SQL = "Select Max([Date]) from [_BS_Material_StatusRecorder]";
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
                    return (DateTime.Now - LastLoadTime).TotalMinutes > 35 || (LastChangeTime - LastLoadTime).Minutes > 0;
                }
            }



            public void reLoad()
            {
                MyRecord.Say("-----初始化物料基本资料开始：------");
                MyRecord.Say("初始化所有纸张/浪纸/物料到本地，请稍等。");
                try
                {
                    string SQL = "Execute [_BS_AllMaterial]";
                    using (MyData.MyDataTable md = new MyData.MyDataTable(SQL))
                    {
                        MyRecord.Say(string.Format("已经读取了{0}条，正在保存到本地。", md.MyRows.Count));
                        reLoad(md.MyRows);
                        LoadTimes += 1;
                        _LastChangeTime = TestLastChangeTime = LastLoadTime = DateTime.Now;
                    }
                }
                catch (Exception ex)
                {
                    MyRecord.Say(ex);
                }
                MyRecord.Say("------初始化物料基本资料结束。-----");
            }

            /// <summary>
            /// 空操作
            /// </summary>
            public void DoNop()
            {
                return;
            }
        }

        /// <summary>
        /// 物料集合
        /// </summary>
        [Serializable]
        public class MaterialCollection : MyBase.MyEnumerable<MaterialItem>
        {
            public MaterialCollection()
            {
            }

            public MaterialCollection(MyData.MyDataRowCollection Data)
            {
                reLoad(Data);
            }

            public MaterialCollection(IEnumerable<MaterialItem> Data)
            {
                reLoad(Data);
            }

            public void reLoad(MyData.MyDataRowCollection Data)
            {
                var v = from a in Data
                        select new MaterialItem(a);
                reLoad(v);
            }

            public void reLoad(IEnumerable<MaterialItem> Data)
            {
                MyRecord.Say("所有物料已经获取，正在保存到本地。");
                _data = Data.ToList();
            }

            public MaterialItem this[string Code]
            {
                get
                {
                    return FindByCode(Code);
                }
                set
                {
                    MaterialItem curItem = FindByCode(Code);
                    if (_current_index >= 0 && _current_index <= Length) curItem = value;
                }
            }

            public MaterialItem FindByCode(string Code)
            {
                MaterialItem rItem = null;
                int i = _data.FindIndex(a => a.Code == Code);
                if (i >= 0)
                {
                    rItem = this[i];
                    _current_index = i;
                }
                return rItem;
            }

            public bool Contains(string Code)
            {
                var v = from a in _data
                        where a.Code == Code
                        select a;
                return (v != null && v.Count() > 0);
            }
        }
        /// <summary>
        /// 物料元素
        /// </summary>
        [Serializable]
        public class MaterialItem : IMaterialItem
        {
            public bool isNew { get; private set; }
            /// <summary>
            /// ID
            /// </summary>
            public int _ID { get; set; }
            /// <summary>
            /// 唯一编号
            /// </summary>
            public Guid CGUID { get; private set; }
            /// <summary>
            /// 类型编号
            /// </summary>
            public string TypeCode { get; set; }
            /// <summary>
            /// 类型
            /// </summary>
            public MaterialTypeItem Type
            {
                get
                {
                    return MaterialTypes[TypeCode];
                }
            }
            /// <summary>
            /// 类型编号
            /// </summary>
            public string CatalogCode
            {
                get
                {
                    return Type != null ? Type.TypeCode : string.Empty;
                }
            }
            /// <summary>
            /// 编号
            /// </summary>
            public string Code { get; private set; }
            /// <summary>
            /// 实际名称
            /// </summary>
            public string PureName { get; set; }
            /// <summary>
            /// 尺寸规格,Specs是这个字段
            /// </summary>
            public string Size { get; set; }
            /// <summary>
            /// 拼接的名称。
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 是否保税_1按照計算的。
            /// </summary>
            public bool Bonded { get; set; }
            /// <summary>
            /// 輸入人
            /// </summary>
            public string Inputer { get; set; }
            /// <summary>
            /// 创建时间
            /// </summary>
            public DateTime InputDate { get; set; }
            /// <summary>
            /// 长
            /// </summary>
            public double Length { get; set; }
            /// <summary>
            /// 宽
            /// </summary>
            public double Width { get; set; }
            /// <summary>
            /// 克重
            /// </summary>
            public double GW { get; set; }
            /// <summary>
            /// 单重
            /// </summary>
            public double Weight { get; set; }
            /// <summary>
            /// 库存上限
            /// </summary>
            public double MaxStock { get; set; }
            /// <summary>
            /// 库存下下
            /// </summary>
            public double MinStock { get; set; }
            /// <summary>
            /// 库存单位
            /// </summary>
            public string Unit { get; set; }
            /// <summary>
            /// 单价单位
            /// </summary>
            public string PriceUnit { get; set; }
            /// <summary>
            /// 单价单位换算比
            /// </summary>
            public double PriceUnitConvertRatio { get; set; }
            /// <summary>
            /// 采购单位
            /// </summary>
            public string PurchaseUnit { get; set; }
            /// <summary>
            /// 采购单位换算比
            /// </summary>
            public double PurchaseUnitConvertRatio { get; set; }

            public double Price { get; set; }

            public double Price1 { get; set; }

            public double Price2 { get; set; }

            public double Price3 { get; set; }

            public double Price4 { get; set; }
            /// <summary>
            /// 采购周期
            /// </summary>
            public double LeadTime { get; set; }
            /// <summary>
            /// 最低起订量
            /// </summary>
            public double MOQ { get; set; }
            /// <summary>
            /// 保质期,天数
            /// </summary>
            public int EXPeriod { get; set; }
            /// <summary>
            /// 呆料标志
            /// </summary>
            public bool Idel { get; set; }
            /// <summary>
            /// 呆料期限,天数
            /// </summary>
            public double IdelPeriod { get; set; }
            /// <summary>
            /// 不計庫存
            /// </summary>
            public bool isNoStock { get; set; }
            /// <summary>
            /// 這是費用
            /// </summary>
            public bool isExpenses { get; set; }
            /// <summary>
            /// 狀態，0-1，正常；>2注銷
            /// </summary>
            public int Status { get; set; }
            /// <summary>
            /// 备注
            /// </summary>
            public string Remark { get; set; }
            /// <summary>
            /// 领料类型
            /// </summary>
            public int PickingType { get; set; }
            /// <summary>
            /// 入库类型
            /// </summary>
            public string InStockType { get; set; }

            /// <summary>
            /// 是否取整
            /// </summary>
            public bool NumberInteger { get; set; }
            /// <summary>
            /// 有效期 天
            /// </summary>
            public double Experiod { get; set; }
            /// <summary>
            /// FSC
            /// </summary>
            public bool FSC { get; set; }

            /// <summary>
            /// 新建一个物料，要给定一个分类。
            /// </summary>
            public MaterialItem(MaterialTypeItem type)
            {
                isNew = true;
                CGUID = Guid.NewGuid();
                TypeCode = type.Code;
            }
            /// <summary>
            /// 平均货存周转天数
            /// </summary>
            public int InventoryTurnOverDays { get; set; }
            /// <summary>
            /// 最后出库时间
            /// </summary>
            public DateTime LastOutStockDate { get; set; }
            /// <summary>
            /// 当前明细库存入库最早时间
            /// </summary>
            public DateTime LotInStockDate { get; set; }

            /// <summary>
            /// 初始化。
            /// </summary>
            /// <param Name="mr"></param>
            public MaterialItem(MyData.MyDataRow mr)
            {
                isNew = false;
                _ID = mr.IntValue("_ID");
                TypeCode = mr.Value("Type");
                Code = mr.Value("Code");
                PureName = mr.Value("PureName");
                Size = mr.Value("Size");
                Name = mr.Value("Name");
                Bonded = mr.BooleanValue("inBond");
                InputDate = mr.Value<DateTime>("InputDate");
                if (InputDate <= DateTime.MinValue) InputDate = new DateTime(2000, 1, 1);
                Length = mr.Value<double>("Length");
                Width = mr.Value<double>("Width");
                GW = mr.Value<double>("gWeight");
                MaxStock = mr.Value<double>("MaxStock");
                MinStock = mr.Value<double>("MinStock");
                Unit = mr.Value("Unit");
                Weight = mr.Value<double>("Weight") * 1000;
                CGUID = mr.GuidValue("CGUID");
                Inputer = mr.Value("Inputer");
                isNoStock = mr.BooleanValue("isNoStock");
                isExpenses = mr.BooleanValue("isExpenses");
                InventoryTurnOverDays = mr.IntValue("InventoryTurnOverDays");
                LastOutStockDate = mr.DateTimeValue("LastOutStockDate");
                LotInStockDate = mr.DateTimeValue("LotInStockDate");
                string x_PriceUnit = mr.Value("PriceUnit");
                if (x_PriceUnit.IsEmptyOrWhiteSpace())
                {
                    PriceUnit = Unit;
                    PriceUnitConvertRatio = 1;
                }
                else
                {
                    PriceUnit = x_PriceUnit;
                    PriceUnitConvertRatio = mr.Value<double>("PriceUnitConvertRatio");
                }
                string x_PurchaseUnit = mr.Value("PurchaseUnit");
                if (x_PurchaseUnit.IsEmptyOrWhiteSpace())
                {
                    PurchaseUnit = Unit;
                    PurchaseUnitConvertRatio = 1;
                }
                else
                {
                    PurchaseUnit = x_PurchaseUnit;
                    PurchaseUnitConvertRatio = mr.Value<double>("PurchaseUnitConvertRatio");
                }
                LeadTime = mr.Value<double>("LeadTime");
                MOQ = mr.Value<double>("MOQ");
                EXPeriod = mr.IntValue("EXPeriod");
                Status = mr.IntValue("Status");
                PickingType = mr.IntValue("PackingType");
                Remark = mr.StringValue("Remark");
                NumberInteger = mr.BooleanValue("NumberInteger");
                FSC = mr.BooleanValue("FSC");
                InStockType = mr.Value("InStockType");
                Price = mr.Value<double>("Price");
                Price1 = mr.Value<double>("Price1");
                Price2 = mr.Value<double>("Price2");
                Price3 = mr.Value<double>("Price3");
                Price4 = mr.Value<double>("Price4");
            }

            public MaterialItem DeepClone()
            {
                using (Stream objectStream = new MemoryStream())
                {
                    IFormatter formatter = new BinaryFormatter();
                    formatter.Serialize(objectStream, (MaterialItem)this);
                    objectStream.Seek(0, SeekOrigin.Begin);
                    return formatter.Deserialize(objectStream) as MaterialItem;
                }
            }

        }


        #endregion
        #region 物料采购价格
        public class MaterialPriceItem
        {
            public MaterialPriceItem(MyData.MyDataRow r)
            {
                Code = Convert.ToString(r["Code"]);
                Price = Convert.ToDouble(r["Price"]);
                MoneyTypeCode = Convert.ToString(r["MoneyTypeCode"]);
                TaxRate = Convert.ToDouble(r["TaxRadio"]);
                ExRate = Convert.ToDouble(r["rate"]);
            }
            public double Price { get; set; }
            public string Code { get; set; }
            public string MoneyTypeCode { get; set; }
            public double TaxRate { get; set; }
            public double ExRate { get; set; }
            /// <summary>
            /// 本幣幣種價格。
            /// </summary>
            public double DefaultMoneyTypePrice
            {
                get
                {
                    MoneyTypeItem mti = MoneyType[MoneyTypeCode];
                    if (mti.IsNotNull())
                        return mti.ToDefaultType(Price);
                    else
                        return 0;
                }
            }

        }
        public class MaterialPriceCollection : MyBase.MyEnumerable<MaterialPriceItem>
        {
            public MaterialPriceItem this[string Code]
            {
                get
                {
                    return this[x => x.Code == Code];
                }
                set
                {
                    this[x => x.Code == Code] = value;
                }
            }
        }

        public class MaterialPrices : MaterialPriceCollection
        {
            public MaterialPrices()
            {
            }

            public void Load()
            {
                string SQL = @"
Select a.MaterialNo as Code,a.Price,c.moneytype as MoneyTypeCode,d.rate,TaxRadio = isNull(a.TaxRadio,c.TaxRadio)
from
pbMaterialProvider a 
Inner Join 
(Select MaterialNo,Max(InputDate) as MaxInputDate from pbMaterialProvider Where isNull(Status,0) > 0 Group by MaterialNo) bb 
ON a.InputDate = bb.MaxInputDate And a.MaterialNo = bb.MaterialNo
Inner Join coVender c On a.VenderNo = c.Code Inner Join pbMoneyType d On c.moneytype = d.Code";
                MyRecord.Say("----加载物料采购价格-----");
                using (MyData.MyDataTable dt = new MyData.MyDataTable(SQL))
                {
                    MyRecord.Say(string.Format("共{0}条", dt.MyRows.Count));
                    var v = from a in dt.MyRows
                            select new MaterialPriceItem(a);
                    this._data = v.ToList();
                }
                MyRecord.Say("----加载物料采购价格，完毕。-----");
            }
        }

        #endregion
        #region 物料采购价格
        private static MaterialPrices _MaterialPrices;
        /// <summary>
        /// 所有物料，纸张物料瓦楞纸
        /// </summary>
        public static MaterialPrices MaterialPrice
        {
            get
            {
                if (_MaterialPrices == null)
                {
                    _MaterialPrices = new MaterialPrices();
                }
                return _MaterialPrices;
            }
        }
        #endregion
        #region 物料分类类

        private static MaterialType _MaterialType;

        public static MaterialType MaterialTypes
        {
            get
            {
                if (_MaterialType == null) _MaterialType = new MaterialType();
                return _MaterialType;
            }
        }

        #endregion
        #region 基本资料的数据逻辑层

        #region 生产工序类

        /// <summary>
        /// 工序
        /// </summary>
        [Serializable]
        public class Process : ProcessCollection
        {
            public Process()
                : base()
            {
                reLoad();
            }

            public void reLoad()
            {
                MyData.MyDataTable _DataSource = null;
                try
                {
                    MyRecord.Say("------初始化工序基本资料------");
                    string SQL = @"Select a.Code as ProcCode,a.name as ProcName,a.pType,WaitTime,a.[_ID] as ID,Remark,otherinf,
                                              FormulaString,FormulaView,Unit,LossRatio,UsePersonFormula,EstimateStatus
                                   From moProcedure a ";
                    _DataSource = new MyData.MyDataTable(SQL);
                    string MemoSQL = @"Select * from _BS_Procedure_Memo", ColorSQL = @"Select code as id,name from pbPrintColor";
                    MyData.MyDataTable MemoData = new MyData.MyDataTable(MemoSQL);
                    MyData.MyDataTable ColorData = new MyData.MyDataTable(ColorSQL);
                    try
                    {
                        MyRecord.Say(string.Format("加载{0}条工序。", _DataSource.MyRows.Count));
                        var q = from a in MemoData.MyRows
                                select new ProcessMemoItem(a.Value("name"), a.Value("ProcCode"), a.IntValue("_ID"))
                                {
                                    InputDate = a.DateTimeValue("InputDate"),
                                    Inputer = a.Value("Inputer"),
                                    SortID = a.Value("SortID")
                                };
                        var qcolor = from a in ColorData.MyRows select new ProcessColorItem(a.IntValue("id"), a.Value("name"));
                        _ProcedureColl.Clear();
                        if (_DataSource != null && _DataSource.Rows.Count > 0)
                        {
                            int dProcessIndex = 0, dProcessCount = _DataSource.MyRows.Count;
                            foreach (var a in _DataSource.MyRows)
                            {
                                dProcessIndex++;
                                string xCode = a.Value("ProcCode");
                                ProcessItem x = new ProcessItem(xCode);
                                string xName = a.Value("ProcName");
                                x.Name = xName;
                                x.pType = a.IntValue("pType");
                                x.WaitTime = a.Value<float>("WaitTime");
                                x.ID = a.IntValue("ID");
                                x.Remark = a.Value("Remark");
                                x.FormulaString = a.Value("FormulaString");
                                x.FormulaView = a.Value("FormulaView");
                                x.LossRatio = a.Value<double>("LossRatio");
                                x.Unit = a.Value("Unit");
                                x.UsePersonFormula = a.BooleanValue("UsePersonFormula");
                                x.isEstimate = !a.BooleanValue("EstimateStatus");
                                if (a.IntValue("pType") == 1)
                                {
                                    x.ColorItemList = qcolor.ToList();
                                }
                                if (q.IsNotEmptySet())
                                {
                                    var vMemoItemList = from u in q
                                                        where u.ProcCode == xCode
                                                        select u;
                                    x.MemoItemList = vMemoItemList.ToList();
                                }
                                else
                                {
                                    x.MemoItemList = null;
                                }
                                _ProcedureColl.Add(xCode, x);
                            }
                            var lv = from a in _ProcedureColl
                                     orderby a.Key
                                     select a.Key;
                            _KeyList = lv.ToList();
                            _LastChangeTime = TestLastChangeTime = LastLoadTime = DateTime.Now;
                            LoadTimes++;
                        }
                        MyRecord.Say("------初始化工序基本资料，完毕------");
                    }
                    finally
                    {
                        MemoData.Dispose();
                        ColorData.Dispose();
                        GC.Collect();
                    }

                }
                finally
                {
                }

            }

            private List<ProcessMemoItem> GenerateMemoItem(string fieldWord, string ProcCode)
            {
                var vWord = fieldWord.Split(',');
                int i = 0;
                var v = from a in vWord
                        select new ProcessMemoItem(a, ProcCode, i++)
                        {
                            SortID = string.Format("{0}{1}", ProcCode, i.ToString().PadLeft(2, '0')),
                            InputDate = DateTime.Now,
                            Inputer = ""
                        };
                return v.ToList();
            }

            private List<ProcessMemoItem> GenerateMemoItem(string ProcCode, List<ProcessMemoItem> MemoItemList)
            {
                var v = from a in MemoItemList
                        where a.ProcCode == ProcCode
                        select a;
                return v.ToList();
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
                        string SQL = "Select Max([Date]) from [_PMC_Process_StatusRecorder]";
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
            /// <summary> 什么都不做
            ///
            /// </summary>
            public void DoNop()
            {

            }

        }


        /// <summary>
        /// 此工序的工艺要求说明
        /// </summary>
        [Serializable]
        public class ProcessMemoItem
        {
            /// <summary>
            /// ID
            /// </summary>
            public int ID { get; private set; }
            /// <summary>
            /// 内容
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 工序编号
            /// </summary>
            public string ProcCode { get; private set; }
            /// <summary>
            /// 排序
            /// </summary>
            public string SortID { get; set; }
            /// <summary>
            /// 输入人
            /// </summary>
            public string Inputer { get; set; }
            /// <summary>
            /// 输入时间
            /// </summary>
            public DateTime InputDate { get; set; }

            public ProcessMemoItem(string ItemWord, string procCode, int _id)
            {
                Name = ItemWord;
                ProcCode = procCode;
                ID = _id;
            }
        }
        /// <summary>
        /// 颜色
        /// </summary>
        [Serializable]
        public class ProcessColorItem
        {
            public ProcessColorItem(int id, string name)
            {
                this.ID = id;
                this.Name = name;
            }
            public int ID { get; private set; }
            public string Name { get; private set; }
            public DateTime InputDate { get; set; }
            public string Inputer { get; set; }
            public void SetName(string text)
            {
                Name = text;
            }
        }

        /// <summary>
        /// 工序类
        /// </summary>
        [Serializable]
        public class ProcessItem
        {
            public MachineCollection Machines { get; private set; }
            /// <summary>
            /// 工序ID
            /// </summary>
            public int ID { get; set; }
            /// <summary>
            /// 工序编号
            /// </summary>
            public string Code { get; set; }
            /// <summary>
            /// 工序名称
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 印前印刷印后分类
            /// </summary>
            public int pType { get; set; }
            /// <summary>
            /// 待干时间
            /// </summary>
            public float WaitTime { get; set; }
            /// <summary>
            /// 备注
            /// </summary>
            public string Remark { get; set; }
            /// <summary>
            /// 工艺要求列表
            /// </summary>
            public List<ProcessMemoItem> MemoItemList { get; set; }

            /// <summary>
            /// 工藝顔色列表
            /// </summary>
            public List<ProcessColorItem> ColorItemList { get; set; }

            public bool isEstimate { get; set; }
            /// <summary>
            /// 是否是印刷
            /// </summary>
            public bool isPrint
            {
                get
                {
                    if (CompanyType =="MY" || CompanyType == "MT" || CompanyType == "MD")
                        return Code == "2000";
                    else
                        return false;
                }
            }
            /// <summary>
            /// 模切工序标志
            /// </summary>
            public bool isDiecut
            {
                get
                {
                    if (CompanyType == "MY" || CompanyType == "MT" || CompanyType == "MD")
                        return Code == "5000";
                    else
                        return false;
                }
            }
            /// <summary>
            /// 书刊装订工序，骑马钉和胶装
            /// </summary>
            public bool isBookbinding
            {
                get
                {
                    if (CompanyType == "MY" || CompanyType == "MT" || CompanyType == "MD")
                        return (Code == "6000" || Code == "6050");
                    else
                        return false;
                }
            }

            public bool isFilm
            {
                get
                {
                    if (CompanyType == "MY" || CompanyType == "MT" )
                        return (Code == "3030");
                    else if (CompanyType == "MD")
                        return (Code == "3060");
                    else
                        return false;
                }
            }



            /// <summary>
            /// 构造函数
            /// </summary>
            /// <param Name="ProcessCode"></param>
            public ProcessItem(string ProcessCode)
            {
                Code = ProcessCode;
                Machines = MainService.Machines.getMachinesByProcess(ProcessCode);
            }

            /// <summary>
            /// 用于估价数量计算的公式
            /// </summary>
            public string FormulaString { get; set; }
            /// <summary>
            /// 公式用于显示的汉字
            /// </summary>
            public string FormulaView { get; set; }
            /// <summary>
            /// 放损率
            /// </summary>
            public double LossRatio { get; set; }
            /// <summary>
            /// 估价的单位
            /// </summary>
            public string Unit { get; set; }
            /// <summary>
            /// 估价是否使用人力计算公式。
            /// </summary>
            public bool UsePersonFormula { get; set; }
        }

        /// <summary>
        /// 所有工序集合。
        /// </summary>
        [Serializable]
        public class ProcessCollection : IEnumerable<ProcessItem>, IEnumerator<ProcessItem>
        {
            #region 存储器
            protected Dictionary<string, ProcessItem> _ProcedureColl = new Dictionary<string, ProcessItem>();
            protected List<string> _KeyList = new List<string>();
            protected int _KeyPoint = -1;

            #endregion 存储器

            #region 构造器

            public ProcessCollection()
            {
            }

            #endregion 构造器

            #region 方法

            public bool Contains(string ProcedureCode)
            {
                return _ProcedureColl != null && ProcedureCode != null && _ProcedureColl.ContainsKey(ProcedureCode);
            }

            #endregion 方法

            #region 索引器

            public ProcessItem this[string key]
            {
                get
                {
                    if (_ProcedureColl.ContainsKey(key))
                    {
                        return _ProcedureColl[key];
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            public ProcessItem this[int index]
            {
                get
                {
                    if (index > -1 && index < _KeyList.Count)
                    {
                        _KeyPoint = index;
                        return (ProcessItem)Current;
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            #endregion 索引器

            #region 枚举接口

            object IEnumerator.Current
            {
                get
                {
                    if (_KeyPoint > -1 && _KeyPoint < _KeyList.Count)
                    {
                        return _ProcedureColl[_KeyList[_KeyPoint]];
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            public ProcessItem Current
            {
                get
                {
                    if (_KeyPoint > -1 && _KeyPoint < _KeyList.Count)
                    {
                        return _ProcedureColl[_KeyList[_KeyPoint]];
                    }
                    else
                    {
                        return null;
                    }
                }
                set
                {
                    if (_KeyPoint > -1 && _KeyPoint < _KeyList.Count)
                    {
                        _ProcedureColl[_KeyList[_KeyPoint]] = value;
                    }
                }
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                _KeyPoint = -1;
                return (IEnumerator)this;
            }

            public IEnumerator<ProcessItem> GetEnumerator()
            {
                _KeyPoint = -1;
                return (IEnumerator<ProcessItem>)this;
            }

            public void Reset()
            {
                _KeyPoint = -1;
            }

            public bool MoveNext()
            {
                _KeyPoint++;
                if (_KeyPoint > -1 && _KeyPoint < _KeyList.Count)
                {
                    return true;
                }
                else
                {
                    _KeyPoint = -1;
                    return false;
                }
            }

            public int Count
            {
                get
                {
                    return _KeyList.Count;
                }
            }

            public void Clear()
            {
                _KeyPoint = -1;
                _KeyList.Clear();
                _ProcedureColl.Clear();
            }

            void IDisposable.Dispose()
            {
            }

            #endregion 枚举接口
        }

        #endregion 工序类

        /// <summary>
        /// 所有工序集合存储器
        /// </summary>
        private static Process _Processes = null;

        /// <summary>
        /// 所有工序
        /// </summary>
        public static Process Processes
        {
            get
            {
                if (_Processes == null)
                    _Processes = new Process();
                else if (_Processes.needReload)
                    _Processes.reLoad();
                return _Processes;
            }
        }


        #endregion
        #region 客户类

        private static Customer _Customer;

        public static Customer Customers
        {
            get
            {
                if (_Customer == null)
                    _Customer = new Customer();
                else if ((DateTime.Now - _Customer.LastLoadTime).TotalMinutes > 75)
                    _Customer.reLoad();
                else if (_Customer.needReload)
                    _Customer.reLoad();
                return _Customer;
            }
        }
        #endregion
        #region 客户类

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
            /// <summary>
            /// 最后加载时间
            /// </summary>
            public DateTime LastLoadTime { get; private set; }
            /// <summary>
            /// 加载次数
            /// </summary>
            public int LoadTimes { get; private set; }

            private DateTime _LastChangeTime = DateTime.Now, TestLastChangeTime = DateTime.Now;

            /// <summary>
            /// 最后修改时间
            /// </summary>
            public DateTime LastChangeTime
            {
                get
                {
                    if ((DateTime.Now - TestLastChangeTime).TotalMinutes > 10)
                    {
                        string SQL = "Select Max([Date]) from [_SD_Customer_StatusRecorder]";
                        object v = MyData.MyCommand.ExecuteScalar(SQL);
                        DateTime oValue;
                        if (DateTime.TryParse(v.ConvertTo<string>(), out oValue))
                        {
                            _LastChangeTime = oValue;
                        }
                        else
                        {
                            _LastChangeTime = DateTime.MinValue;
                        }
                        TestLastChangeTime = DateTime.Now;
                    }
                    return _LastChangeTime;
                }
            }
            /// <summary>
            /// 需要重新加载
            /// </summary>
            public bool needReload
            {
                get
                {
                    return (DateTime.Now - LastLoadTime).TotalMinutes > 90 || (LastChangeTime - LastLoadTime).Minutes > 0;
                }
            }
            public void reLoad()
            {
                string SQL = @"
Select a.*,
       b.cname as MoneyTypeName,c.name as TypeName,d.Name as DeliveryTypeName,e.Name as PaymentTypeName
from coCustomer a Left Outer Join pbMoneyType b ON a.MoneyType = b.Code 
                  Left Outer Join coClientType c ON a.Type = c.Code 
                  Left Outer Join coSendType d ON a.SendType=d.Code
                  Left Outer Join coPayType e ON a.PayTime=e.Code
";
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL))
                {
                    reLoad(md.MyRows);
                    LoadTimes += 1;
                    _LastChangeTime = TestLastChangeTime = LastLoadTime = DateTime.Now;
                }
            }

            static public new void Clear()
            {
                _Customer = null;
            }

            public void DoNop()
            {
                return;
            }

            static public bool IsAppleCustomer(string CustCode)
            {
                if (CompanyType == "MY"|| CompanyType == "ZS")
                    return (CustCode == "S062");
                else if (CompanyType == "MT")
                {
                    string[] AppleCustCode = new string[] { "YM001", "YX005", "YX025" };
                    string CustCodeTest = CustCode.Split("#").FirstOrDefault();
                    return AppleCustCode.Contains(CustCodeTest);
                }
                else if (CompanyType == "MD")
                {
                    return (CustCode == "SA001");
                }
                else
                {
                    return (CustCode == "S062");
                }
            }

            public static string CustomerCodeRegexExpression
            {
                get
                {
                    try
                    {
                        if (CompanyType == "MY")
                            return @"(S[0-9]{3})|(MY)|(Y[MX][0-9]{3,4}\#?)";
                        else if (CompanyType == "MT")
                            return @"(S[0-9]{3})|(MY)|(Y[MX][0-9]{3,4}\#?)";
                        else if (CompanyType == "ZS")
                            return @"(S[0-9]{3})|(MY)|(Y[MX][0-9]{3,4}\#?)";
                        else if (CompanyType == "MD")
                            return @"([A-Za-z\d]{0,5})";
                        else
                            return "";
                    }
                    catch (Exception ex)
                    {
                        return "";
                    }
                }
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

            public CustomerCollection(MyData.MyDataRowCollection CustomerData)
            {
                reLoad(CustomerData);
            }

            public CustomerCollection(IEnumerable<CustomerItem> Customers)
            {
                reLoad(Customers);
            }

            public void reLoad(MyData.MyDataRowCollection CustomerData)
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
                    id = m.Value<int>("_id");
                    xid = m.Value<int>("id");
                    TypeCode = m.Value<string>("Type");
                    TypeName = m.Value<string>("TypeName");
                    Name = m.Value<string>("Name");
                    Code = m.Value<string>("Code").Trim().ToUpper();
                    Address = m.Value<string>("SendAddress");
                    if (Address == null || Address.Length <= 0) Address = m.Value("Address");
                    SalesMan = m.Value("BusinessMan");
                    Assistant = m.Value("Assistant");
                    OEMCustomer = m.Value("OEMCustomer");
                    CGUID = m.Value<Guid>("CGUID");
                    CreateDate = m.Value<DateTime>("CreateDate");
                    if (CreateDate <= DateTime.MinValue) CreateDate = m.Value<DateTime>("setDate");
                    Creator = m.Value("Creator");
                    setDate = m.Value<DateTime>("setDate");
                    PaymentTypeCode = m.Value("paytime");
                    PaymentTypeName = m.Value("PaymentTypeName");
                    MoneyTypeCode = m.Value("moneytype");
                    MoneyTypeName = m.Value("MoneyTypeName");
                    TelNumber = m.Value("tel");
                    FaxNumber = m.Value("fax");
                    Connector = m.Value("connecter");
                    IntelAddress = m.Value("inteladdress");
                    EMailAddress = m.Value("email");
                    Status = m.IntValue("status");
                    Remark = m.StringValue("remark");
                    DeliveryTypeCode = m.Value("SendType");
                    DeliveryTypeName = m.Value("DeliveryTypeName");
                    DepartmentCode = m.Value("DepartmentCode");
                    DeliveryNoteWeightPrecision = m.Value<int>("DeliveryNoteWeightPrecision");
                }
            }

            public CustomerItem(string TypeCode)
            {
                this.CreateDate = DateTime.Now;
                this.Creator = "";
                this.TypeCode = TypeCode;
                CGUID = Guid.NewGuid();
                isNew = true;
            }

            public void LoadSendAddressList()
            {
                string SQL = "Select * from [_SD_Customer_SendAddress] Where Code = @Code                 ";
                using (MyData.MyDataTable SendAddressDataSource = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@Code", this.Code)))
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
            /// 送货单单重小数位
            /// </summary>
            public int DeliveryNoteWeightPrecision { get; set; }

            public bool isApple
            {
                get
                {
                    return Customer.IsAppleCustomer(this.Code);
                }
            }
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

        #endregion 客户类
        #region 工单/BOM类

        #region 工单/BOM部件
        /// <summary> ＢＯＭ部件/工单部件
        ///
        /// </summary>
        [Serializable]
        public class PartItem
        {
            [Serializable]
            public enum PartAssmeblyType
            {
                /// <summary>
                /// 物料领料
                /// </summary>
                Material = 100,
                /// <summary>
                /// 部件加工
                /// </summary>
                Part = 200,
                /// <summary>
                /// 成品加工
                /// </summary>
                Assmebly = 300,
                /// <summary>
                /// 成品入库
                /// </summary>
                Product = 500
            }

            public PartItem(string partid)
            {
                PartID = partid;
            }

            public PartItem()
            {
            }
            /// <summary>
            /// 部件名称
            /// </summary>
            public string PartID { get; set; }
            /// <summary>
            /// 印刷部件名
            /// </summary>
            public string PrintPartID
            {
                get
                {
                    if (PartID.Length > 0)
                    {
                        Regex rgx = new Regex(@"T\d{2}-\d{2}");
                        if (rgx.IsMatch(PartID))
                            return PartID.Substring(1, 3);
                        else
                            return PartID;
                    }
                    else
                        return string.Empty;
                }
            }
            /// <summary>
            /// 部件的ID，用于生产工序和物料ID
            /// </summary>
            public int PartIndex
            {
                get
                {
                    if (PartID.Length > 0)
                    {
                        Regex rgx = new Regex(@"[tT]\d{1,2}-\d{1,2}");
                        if (rgx.IsMatch(PartID))
                        {
                            rgx = new Regex(@"(?<=[Tt])\d{1,2}");
                            string atxt = rgx.Match(PartID).Value;
                            int a = 0; if (!int.TryParse(atxt, out a)) a = 0;
                            a = (a + 1) * 100;
                            rgx = new Regex(@"(?<=-)\d{1,2}$");
                            string btxt = rgx.Match(PartID).Value;
                            int b = 0; if (!int.TryParse(btxt, out b)) b = 0;
                            return (a + b) * 100;
                        }
                        else
                        {
                            rgx = new Regex(@"[tT]\d{2}");
                            if (rgx.IsMatch(PartID))
                            {
                                rgx = new Regex(@"(?<=[Tt])\d{1,2}");
                                string atxt = rgx.Match(PartID).Value;
                                int a = 0; if (!int.TryParse(atxt, out a)) a = 0;
                                a = (a + 1) * 1000;
                                return a;
                            }
                            else
                            {
                                return 800000;
                            }
                        }
                    }
                    else
                    {
                        return 900000;
                    }
                }
            }
            /// <summary>
            /// 部件加工还是成品加工
            /// </summary>
            public PartAssmeblyType PartType
            {
                get
                {
                    if (PartID.Length > 0)
                        return PartAssmeblyType.Assmebly;
                    else
                        return PartAssmeblyType.Part;
                }
            }

        }
        #endregion

        #region 工单类

        public static BomMaterial.BomMaterialTypeEnum GetMaterialBomType(string Catelog, string Type, string Name)
        {
            int cat = Catelog.ConvertTo<int>(0, true);
            if (cat == 1)
                return BomMaterial.BomMaterialTypeEnum.BomProduct;
            else if (cat == 2)
                return BomMaterial.BomMaterialTypeEnum.BomPaper;
            else if (cat == 3)
                return BomMaterial.BomMaterialTypeEnum.BomWavePaper;
            else if (cat == 4)
            {
                if (Name.Contains("APET"))
                    return BomMaterial.BomMaterialTypeEnum.BomOPP;
                else if (Type == "403")
                    return BomMaterial.BomMaterialTypeEnum.BomINK;
                else
                    return BomMaterial.BomMaterialTypeEnum.BomOtherMaterial;
            }
            else
                return BomMaterial.BomMaterialTypeEnum.BomOtherMaterial;
        }

        /// <summary> 工单类
        ///
        /// </summary>
        [Serializable]
        public class ProduceNote : ProduceNoteItem
        {
            #region 构造器

            public ProduceNote(string ProduceRdsNo)
            {
                MyRecord.Say(string.Format("{0}，进入工单类。", ProduceRdsNo));
                Load(ProduceRdsNo);
                LoadOrderInfo();
                LoadDetial();
                isNew = false;
                isEdit = false;
                MyRecord.Say(string.Format("{0}，工单读取完毕。", ProduceRdsNo));
            }

            public ProduceNote(int ProduceID)
                : base()
            {
                ID = ProduceID;
            }

            public void ReLoad()
            {
                if (isNew) return;
                if (this.ID != 0)
                {
                    Load(this.ID);
                    LoadOrderInfo();
                    LoadDetial();
                }
            }

            private void Load(string RdsNo)
            {
                MyRecord.Say(string.Format("{0}，读取工单基本信息。", RdsNo));
                string SQL = @"Select * from moProduce Where rdsno = @RdsNo";
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@RdsNo", RdsNo)))
                {
                    if (md.Rows.Count > 0)
                    {
                        MyData.MyDataRow xRow = md.MyRows.SingleOrDefault();
                        FillData(xRow);
                    }
                }
            }

            private void Load(int ProduceID)
            {
                if (isNew) return;
                if (ProduceID > 0)
                {
                    string SQL = @"Select * from moProduce Where ID = @id";
                    using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@id", ProduceID, MyData.MyDataParameter.MyDataType.Int)))
                    {
                        if (md.Rows.Count > 0)
                        {
                            MyData.MyDataRow xRow = md.MyRows.SingleOrDefault();
                            FillData(xRow);
                        }
                    }
                }
            }

            public void Load()
            {
                if (isNew) return;
                if (this.ID != 0)
                {
                    Load(this.ID);
                }
            }

            private void FillData(MyData.MyDataRow curRow)
            {
                if (curRow != null)
                {
                    this.ID = Convert.ToInt32(curRow["ID"]);
                    this.RdsNo = Convert.ToString(curRow["RdsNo"]);
                    this.ProdCode = Convert.ToString(curRow["Code"]);
                    this.OrderNo = Convert.ToString(curRow["OrderNo"]);
                    this.CustCode = Convert.ToString(curRow["custid"]);
                    this.PNumb = Convert.ToInt32(curRow["pnumb"]);
                    this.DeliveryDate = Convert.ToDateTime(curRow["pDeliver"]);
                    this.OrderDeliveryDate = Convert.ToDateTime(curRow["oDeliver"]);
                    this.InputDate = Convert.ToDateTime(curRow["InputDate"]);
                    this.Inputer = Convert.ToString(curRow["Inputer"]);
                    this.CheckDate = Convert.ToDateTime(curRow["CheckDate"]);
                    this.Checker = Convert.ToString(curRow["Checker"]);
                    this.Remark = Convert.ToString(curRow["Remark"]);
                    this.OrderSID = Convert.ToInt32(curRow["PID"]);
                    this.RemarkChange = Convert.ToString(curRow["RemarkChange"]);
                    this.RemarkReDo = Convert.ToString(curRow["RemarkRenew"]);
                    this.Status = Convert.ToInt32(curRow["Status"]);
                    this.StockStatus = Convert.ToInt32(curRow["StockStatus"]);
                    this.BeforeStatusID = Convert.ToInt32(curRow["BeforeStatusID"]);
                    this.PropertyID = Convert.ToInt32(curRow["Property"]);
                    this.isMutipart = Convert.ToBoolean(curRow["mPart"]);
                    this.OriginalRdsNo = Convert.ToString(curRow["RenewRdsNo"]);
                    this.FinishDate = Convert.ToDateTime(curRow["finishdate"]);
                    this.StockFinishDate = Convert.ToDateTime(curRow["StockDate"]);
                    this.VerType = Convert.ToString(curRow["VerType"]);
                    this.isStopStock = (Convert.ToInt32(curRow["StopStock"]) == 1);
                    this.LastProcessCode = Convert.ToString(curRow["prodprocid"]);
                    this.LastProcessID = Convert.ToInt32(curRow["LastProcessID"]);
                    this.FinishMan = Convert.ToString(curRow["FinishMan"]);
                    this.FinishRemark = Convert.ToString(curRow["FinishRemark"]);
                    this.DeliveryCheckDate = Convert.ToDateTime(curRow["RTime"]);
                    this.DeliveryChecker = Convert.ToString(curRow["RMan"]);
                    this.DeliveryRemark = Convert.ToString(curRow["SRemark"]);
                    this.RejectRdsNo = Convert.ToString(curRow["RejectRdsNo"]);
                    this.PlateCode = Convert.ToString(curRow["PlateCode"]);
                    this.CutPlateCode = Convert.ToString(curRow["dm"]);
                    this.HotPlateCode = Convert.ToString(curRow["HotPlateCode"]);
                    this.ScreenPlateCode = Convert.ToString(curRow["ScreenPlateCode"]);
                    this.UseStockNumb = Convert.ToSingle(curRow["snumb"]);
                    this.ComplainRdsNo = Convert.ToString(curRow["ComplainRdsNo"]);
                    this.ReturnRdsNo = Convert.ToString(curRow["ReturnRdsNo"]);
                    this.isBonded = Convert.ToInt32(curRow["Bonded"]) == 1;
                    this.InStockLocationType = Convert.ToString(curRow["StockLocationType"]);
                }
            }

            #region 加载订单信息
            public void LoadOrderInfo()
            {
                MyRecord.Say(string.Format("{0}，读取工单SO信息。", RdsNo));
                if (isNew) return;
                Thread.Sleep(200);
                string SQL = @" Select a.id as Mid,custorderid as CustOrderNo,a.Remark,b.Remark as sRemark,b.PayNumb as Numb
                                from coOrder a,coOrderProd b Where a.id=b.zbid And a.RdsNo=@OrderNo And b.id=@PID ";
                MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
                {
                    new MyData.MyDataParameter("@OrderNo",OrderNo),
                    new MyData.MyDataParameter("@PID",OrderSID, MyData.MyDataParameter.MyDataType.Int)
                };
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, mps))
                {
                    if (md != null && md.Rows.Count > 0)
                    {
                        MyData.MyDataRow mr = md.MyRows.FirstOrOnlyRow;
                        _OrderMid = Convert.ToInt32(mr["Mid"]);
                        _OrderNumb = Convert.ToInt32(mr["Numb"]);
                        _OrderRemark = string.Format("{0}\r\n{1}", Convert.ToString(mr["remark"]), Convert.ToString(mr["sRemark"]));
                        _CustOrderNo = Convert.ToString(mr["CustOrderNo"]);
                    }
                }
            }
            #endregion

            #region 加载领料数
            /// <summary>
            /// 加载领料数
            /// </summary>
            public void LoadPickNumb()
            {
                if (InputDate <= Convert.ToDateTime("2015-05-20 00:00:00") || CompanyType == "MY" )
                {
                    MyRecord.Say("加载工单的领料数。");
                    if (this.isMutipart)
                    {
                        string SQL = @"Select a.Numb,isNull(a.yCode,a.Code) as Code from stOutProdLst a Where ProductNo=@RdsNo";
                        string SQLB = @"Select a.Numb,a.Code,a.PartID from stOtherOutLst a,stOtherOut b Where a.zbid=b.id And b.OtherNo=@RdsNo";
                        MyData.MyDataParameter mp = new MyData.MyDataParameter("@RdsNo", this.RdsNo);
                        MyData.MyDataTable md = new MyData.MyDataTable(SQL, mp);
                        MyData.MyDataTable mb = new MyData.MyDataTable(SQLB, mp);
                        foreach (var item in this.Materials)
                        {
                            var vx1 = from a in md.MyRows
                                      where Convert.ToString(a["Code"]) == item.Code
                                      select Convert.ToDouble(a["Numb"]);
                            var vx2 = from a in mb.MyRows
                                      where Convert.ToString(a["PartID"]) == item.PartID
                                      select Convert.ToDouble(a["Numb"]);
                            double p1 = vx1.Sum();
                            if (p1 > Convert.ToDouble(item.Number))
                            {
                                item.PickedNumber = Convert.ToDouble(item.Number);
                                item.OverflowNumber = vx2.Sum();
                            }
                            else
                            {
                                item.PickedNumber = p1;
                                item.OverflowNumber = vx2.Sum();
                            }
                        }
                    }
                    else
                    {
                        string SQL = @"Select a.Numb,isNull(a.yCode,a.Code) as Code from stOutProdLst a Where ProductNo=@RdsNo";
                        string SQLB = @"Select a.Numb,a.Code from stOtherOutLst a,stOtherOut b Where a.zbid=b.id And b.OtherNo=@RdsNo";
                        MyData.MyDataParameter mp = new MyData.MyDataParameter("@RdsNo", this.RdsNo);
                        MyData.MyDataTable md = new MyData.MyDataTable(SQL, mp);
                        MyData.MyDataTable mb = new MyData.MyDataTable(SQLB, mp);
                        foreach (var item in this.Materials)
                        {
                            var vx1 = from a in md.MyRows
                                      where Convert.ToString(a["Code"]) == item.Code
                                      select Convert.ToDouble(a["Numb"]);
                            var vxx1 = from a in md.MyRows
                                       where Convert.ToString(a["Code"]).Substring(0, 1) == item.Code.Substring(0, 1)
                                       group a by Convert.ToString(a["Code"]) into g
                                       select g.Key;
                            var vx2 = from a in mb.MyRows
                                      where Convert.ToString(a["Code"]) == item.Code
                                      select Convert.ToDouble(a["Numb"]);
                            if (vxx1.Count() == 1)
                            {
                                var vx3 = from a in mb.MyRows
                                          where Convert.ToString(a["Code"]) != item.Code && Convert.ToString(a["Code"]).Substring(0, 1) == item.Code.Substring(0, 1) && item.Code.Substring(0, 1) == "2"
                                          select Convert.ToDouble(a["Numb"]);
                                item.PickedNumber = vx1.Sum();
                                item.OverflowNumber = vx2.Sum() + vx3.Sum();
                            }
                            else
                            {
                                item.PickedNumber = vx1.Sum();
                                item.OverflowNumber = vx2.Sum();
                            }
                        }

                    }
                }
            }

            #endregion


            #endregion 构造器

            #region 覆写的属性
            #region 订单表内容
            /// <summary>
            /// 订单ID
            /// </summary>
            public override int OrderMID
            {
                get
                {
                    return _OrderMid;
                }
                set { _OrderMid = value; }
            }

            /// <summary>
            /// 订单总数
            /// </summary>
            public override int OrderNumb
            {
                get
                {
                    return _OrderNumb;
                }
                set { _OrderNumb = value; }
            }

            /// <summary>
            /// 客户单号
            /// </summary>
            public override string CustOrderNo
            {
                get
                {
                    return _CustOrderNo;
                }
                set { _CustOrderNo = value; }
            }

            /// <summary>
            /// 订单备注
            /// </summary>
            public override string OrderRemark
            {
                get
                {
                    return _OrderRemark;
                }
                set { _OrderRemark = value; }
            }
            #endregion
            #endregion

            #region 集合属性

            public void LoadDetial()
            {
                if (isNew) return;
                MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
                {
                    new MyData.MyDataParameter("@RdsNo",RdsNo),
                    new MyData.MyDataParameter("@ID",ID, MyData.MyDataParameter.MyDataType.Int),
                };
                #region 加载工艺
                MyRecord.Say(string.Format("{0}，读取工单工艺。", RdsNo));
                string SQLProcess = @"Exec [_PMC_ProduceProcess] @ID";
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQLProcess, mps))
                {
                    var vp = from a in md.MyRows
                             select new ProduceProcess(a, this);
                    _ProduceProcesses = vp.ToList();
                }
                #endregion
                #region 加载纸张物料
                MyRecord.Say(string.Format("{0}，读取工单物料。", RdsNo));
                string SQLMaterial = @"Exec [_PMC_ProduceMaterial] @ID";
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQLMaterial, mps))
                {
                    var vm = from a in md.MyRows
                             select new ProduceMaterial(a, this);
                    _ProduceMaterials = vm.ToList();
                }
                #endregion
                #region 加载部件
                MyRecord.Say(string.Format("{0}，生成工单明细信息。", RdsNo));
                if (_ProduceProcesses.Count + _ProduceMaterials.Count > 0)
                {
                    if (_ProduceMaterials != null && _ProduceProcesses != null)
                    {
                        var vPart = from a in
                                        (from xa in _ProduceProcesses select xa.PartID).Union(from xb in _ProduceMaterials select xb.PartID)
                                    where a.Length > 0
                                    group a by a into g
                                    select new PartItem(g.Key);
                        _ProduceParts = vPart.ToList();
                    }
                    else if (_ProduceProcesses != null)
                    {
                        var vPart = from a in
                                        (from xa in _ProduceProcesses select xa.PartID)
                                    where a.Length > 0
                                    group a by a into g
                                    select new PartItem(g.Key);
                        _ProduceParts = vPart.ToList();
                    }
                    else if (_ProduceMaterials != null)
                    {
                        var vPart = from a in
                                        (from xb in _ProduceMaterials select xb.PartID)
                                    where a.Length > 0
                                    group a by a into g
                                    select new PartItem(g.Key);
                        _ProduceParts = vPart.ToList();
                    }
                    _isMutiPart = (_ProduceParts.Count > 0);

                    //if (!_isMutiPart)
                    //{
                    //    if (_ProduceMaterials != null) _ProduceMaterials.ForEach(new Action<ProduceMaterial>(x => x.PartID = string.Empty));
                    //    if (_ProduceProcesses != null) _ProduceProcesses.ForEach(new Action<ProduceProcess>(x => x.PartID = string.Empty));
                    //}
                }
                #endregion
            }



            /// <summary>
            /// 加载拆解入库物料表
            /// </summary>
            public void LoadDestoryProduct()
            {
                if (PropertyID == 5)
                {
                    MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
                    {
                        new MyData.MyDataParameter("@RdsNo",RdsNo),
                        new MyData.MyDataParameter("@ID",ID, MyData.MyDataParameter.MyDataType.Int),
                    };
                    string SQLDestory = "Select a.*,Name=b.FullName,b.Unit,b.Size from [_PMC_Produce_DestroyProduct] a Inner Join [AllMaterialView] b on a.Code=b.Code Where zbid=@ID";
                    using (MyData.MyDataTable md = new MyData.MyDataTable(SQLDestory, 20, mps))
                    {
                        if (md != null && md.MyRows.Count > 0)
                        {
                            var vdestory = from a in md.MyRows
                                           select new ProduceDestroyProduct()
                                           {
                                               ID = Convert.ToInt32(a["id"]),
                                               zbid = Convert.ToInt32(a["zbid"]),
                                               Code = Convert.ToString(a["code"]),
                                               Name = Convert.ToString(a["Name"]),
                                               Numb = Convert.ToDouble(a["Numb"]),
                                               UnitNumb = Convert.ToDouble(a["UnitNumb"]),
                                               Remark = Convert.ToString(a["Remark"])
                                           };
                            _ProduceDestroyProduct = vdestory.ToList();
                        }
                    }
                }
            }

            #region 集合属性
            /// <summary>
            /// 部件
            /// </summary>
            public override List<PartItem> Parts
            {
                get
                {
                    return _ProduceParts;
                }
                set { _ProduceParts = value; }
            }
            /// <summary>
            /// 物料集合
            /// </summary>
            public override List<ProduceMaterial> Materials
            {
                get
                {
                    return _ProduceMaterials;
                }
                set { _ProduceMaterials = value; }
            }
            /// <summary>
            /// 工序集合
            /// </summary>
            public override List<ProduceProcess> Processes
            {
                get
                {
                    return _ProduceProcesses;
                }
                set { _ProduceProcesses = value; }
            }
            #endregion

            #endregion

            #region 操作
            /// <summary>
            /// 设置补单
            /// </summary>
            public virtual void MakeReDo()
            {
                OriginalRdsNo = RdsNo;
                RdsNo = string.Empty;
                ID = 0;
                PropertyID = 1;
                Status = 0;
                StockFinishDate = DateTime.MinValue;
                FinishDate = DateTime.MinValue;
                FinishMan = null;
                FinishRemark = null;
                StockStatus = 0;
                isNew = true;
                isEdit = true;
            }

            /// <summary>
            /// 获取最终工序
            /// </summary>
            /// <returns></returns>
            public virtual ProduceProcess getLastProcessCode()
            {
                if (Processes != null)
                {
                    var v = from a in Processes
                            orderby a.ID descending
                            select a;
                    if (v.Count() > 0)
                    {
                        LastProcessCode = v.FirstOrDefault().Process.Code;
                        LastProcessID = v.Max(a => a.ID);
                        return v.FirstOrDefault();
                    }
                }
                return null;
            }
            /// <summary>
            /// 计算所有，包括应产数和物料
            /// </summary>
            public virtual void CalcateNumber()
            {
                CalculateProcessPNumb();
                CalculateMaterialFromProcess();
            }
            /// <summary>
            /// 计算各工序应产数
            /// </summary>
            public virtual void CalculateProcessPNumb()
            {
                if (Processes != null && Processes.Count > 0)
                {
                    var q = from a in Processes
                            where a.PartID.Length <= 0
                            orderby a.ID descending
                            select a;
                    if (q.Count() > 0)
                    {
                        ProduceProcess pp = q.FirstOrDefault();
                        pp.PNumb = PNumb / pp.ColNumb;
                        pp.CalcuatePreviousProcess();
                    }
                    else
                    {
                        var qx = from a in Processes
                                 where a.PartID.Length >= 0
                                 orderby a.ID descending
                                 group a by a.PartID into g
                                 select g.FirstOrDefault();
                        if (qx.Count() > 0)
                        {
                            foreach (var qitem in qx)
                            {
                                qitem.PNumb = PNumb / qitem.ColNumb;
                                qitem.CalcuatePreviousProcess();
                            }
                        }
                    }
                }
            }
            /// <summary>
            /// 根据应产数算物料
            /// </summary>
            public virtual void CalculateMaterialFromProcess()
            {
                ///根据应产数算物料
                if (Materials != null && Materials.Count > 0)
                {
                    foreach (var item in Materials)
                    {
                        item.PartNumb = PNumb;
                        string xPart = item.PartID;
                        if (item.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper)
                        {
                            var pp = from a in Processes
                                     where (a.PartID.Length <= 0 || a.PartID == xPart) && a.Process.pType == 1
                                     select a.WastageNumb * a.ColNumb;
                            item.PrintLossNumb = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(pp.Sum()) / Convert.ToDouble(item.ColNumb)));

                            var pv = from a in Processes
                                     where (a.PartID.Length <= 0 || a.PartID == xPart) && a.Process.pType > 1
                                     select a.WastageNumb * a.ColNumb;
                            item.PostpressLossNumb = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(pv.Sum()) / Convert.ToDouble(item.ColNumb)));

                            var pw = from a in Processes
                                     where (a.PartID.Length <= 0 || a.PartID == xPart) && a.Process.pType > 1
                                     select a.UseRemainingNumb;
                            item.UseStockNumb = pw.Sum();
                        }
                        else if (item.MaterialType == BomMaterial.BomMaterialTypeEnum.BomWavePaper)
                        {
                            int procid = item.ProcID;
                            var ww = from a in Processes
                                     where (a.PartID.Length <= 0 || a.PartID == xPart) && a.ID >= procid
                                     select a.WastageNumb * a.ColNumb;
                            var wp = from a in Materials
                                     where (a.PartID.Length <= 0 || a.PartID == xPart) && a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper
                                     select a.PrintLossNumb * a.ColNumb;
                            if (ww.Count() > 0)
                            {
                                double sum1 = Convert.ToDouble(ww.Sum()) / Convert.ToDouble(item.ColNumb), sum2 = Convert.ToDouble(wp.Sum()) / Convert.ToDouble(item.ColNumb) * 0.15d;
                                item.PostpressLossNumb = Convert.ToDecimal(Math.Ceiling(sum1 + sum2));
                            }
                        }
                        else if (item.MaterialType == BomMaterial.BomMaterialTypeEnum.BomINK)
                        {
                            int procid = item.ProcID;
                            var ww = from a in Processes
                                     where (a.PartID.Length <= 0 || a.PartID == xPart) && a.ID == procid
                                     select a.PNumb + a.WastageNumb;
                            if (ww.Count() > 0)
                            {
                                item.PartNumb = ww.Sum();
                            }
                        }
                        else
                        {
                            int procid = item.ProcID;
                            var ww = from a in Processes
                                     where (a.PartID.Length <= 0 || a.PartID == xPart) && a.ID >= procid
                                     select a.WastageNumb * a.ColNumb;
                            if (ww.Count() > 0)
                            {
                                double sum1 = Convert.ToDouble(ww.Sum());
                                item.PostpressLossNumb = Convert.ToDecimal(Math.Ceiling(sum1));
                            }
                        }
                    }
                }

                if (PropertyID == 5)
                {
                    if (DestroyProduct != null && DestroyProduct.Count > 0)
                    {
                        foreach (var item in DestroyProduct)
                        {
                            item.Numb = PNumb * item.UnitNumb;
                        }
                    }
                }
            }
            /// <summary>
            /// 简单计算物料领料数
            /// </summary>
            public virtual void CalcateMaterialNumbLimite()
            {
                if (Materials != null && Materials.Count > 0)
                {
                    foreach (var item in Materials)
                    {
                        if (item.MaterialType == BomMaterial.BomMaterialTypeEnum.BomWavePaper || item.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper)
                        {
                            item.Number = Math.Ceiling(Convert.ToDouble(item.SheetNumber + item.PrintLossNumb + item.PostpressLossNumb) / Convert.ToDouble(item.CutNumb));
                        }
                        else
                        {
                            if (item.Material.IsNotNull() && item.Material.NumberInteger) //如果取整。
                                item.Number = Math.Ceiling(Math.Round(Convert.ToDouble(item.UnitNumb) * Convert.ToDouble(item.PartNumb + item.PostpressLossNumb), 1));
                            else
                                item.Number = Convert.ToDouble(item.UnitNumb) * Convert.ToDouble(item.PartNumb + item.PostpressLossNumb);
                        }
                    }
                }
            }

            /// <summary>
            /// 更新状态和各种数字。
            /// </summary>
            /// <returns></returns>
            public bool UpdatePickingAndFinishNumb()
            {
                MyData.MyCommand md = new MyData.MyCommand();
                if (RdsNo != null && RdsNo.Length > 0)
                {
                    string SQL = @"Execute [_PMC_UpdateFinishAndPicking] @rdsno";
                    MyData.MyDataParameter mp = new MyData.MyDataParameter("@rdsno", RdsNo);
                    return md.Execute(SQL, mp);
                }
                else
                {
                    string SQL = @"
                        Declare @rdsno Varchar(100)
                        Set @rdsno=(Select RdsNo from moProduce Where id=@id)
                        Execute [_PMC_UpdateFinishAndPicking] @rdsno";
                    MyData.MyDataParameter mp = new MyData.MyDataParameter("@id", ID, MyData.MyDataParameter.MyDataType.Int);
                    return md.Execute(SQL, mp);
                }
            }


            #endregion

            #region 入库数量
            /// <summary>
            /// 生产完工数量
            /// </summary>
            public int FinishNumber
            {
                get
                {
                    var v = from a in this.Processes
                            where a.NextProcess == null
                            orderby a.ID descending
                            select a.FinishNumb;
                    if (v != null && v.Count() > 0)
                    {
                        return Convert.ToInt32(v.FirstOrDefault());
                    }
                    return 0;
                }
            }

            /// <summary>
            /// 入库数
            /// </summary>
            public int StockNumber
            {
                get
                {
                    string SQL = string.Empty;
                    if (this.PropertyID == 5)
                        SQL = "Select Max(Numb) as Numb From stprdstocklst b Where b.ProductNo= @RdsNo";
                    else
                        SQL = "Select Sum(Numb) as Numb From (SELECT Sum(NUMB) as Numb,Code,ProductNo FROM stPrdStocklst Group by CODE,ProductNo) m Where m.ProductNo= @RdsNo";

                    MyData.MyDataParameter mp = new MyData.MyDataParameter("@RdsNo", this.RdsNo);
                    MyData.MyDataTable md = new MyData.MyDataTable(SQL, mp);
                    if (md != null && md.MyRows.Count > 0)
                        return md.MyRows.FirstOrOnlyRow.IntValue("Numb"); //Convert.ToInt32(md.MyRows.FirstOrDefault()["Numb"]);
                    else
                        return 0;
                }
            }

            public int WIPOrderCount
            {
                get
                {
                    string SQL = @"Select Count(*) as N from moProduce a 
                                    Where isNull(a.status,0)>0 And ((isNull(a.StopStock,0)=0 And a.StockDate is Null) or (isNull(a.StopStock,0)=1 and FinishDate is Null))
                                      AND a.Code = @Code";
                    MyData.MyDataParameter mp = new MyData.MyDataParameter("@Code", this.ProdCode);
                    MyData.MyDataTable md = new MyData.MyDataTable(SQL, mp);
                    if (md.MyRows.IsNotEmptySet())
                    {
                        return md.MyRows.FirstOrOnlyRow.IntValue("N");
                    }
                    else
                    {
                        return 0;
                    }
                }
            }


            #endregion
        }

        [Serializable]
        public class ProduceNoteItem
        {
            public ProduceNoteItem()
            {
                isNew = false;
                isEdit = false;
            }

            #region 属性存储器

            public virtual bool isNew { get; protected set; }

            public virtual bool isEdit { get; set; }

            protected bool _isMutiPart = false;
            /// <summary>
            /// 是否多部件
            /// </summary>
            public virtual bool isMutipart { get { return _isMutiPart; } set { _isMutiPart = value; } }

            public virtual bool isStopStock { get; set; }

            /// <summary>
            /// 工单号
            /// </summary>
            public virtual string RdsNo { get; protected set; }

            /// <summary>
            /// 工单ID
            /// </summary>
            public virtual int ID { get; protected set; }

            protected string _ProdCode;
            /// <summary>
            /// 产品编号
            /// </summary>
            public virtual string ProdCode
            {
                get { return _ProdCode; }
                protected set
                {
                    _ProdCode = value;
                    if (!Products.Contains(value)) Products.reLoad();
                    if (Products.Contains(value)) Product = Products[value];
                }
            }
            /// <summary>
            /// 产品
            /// </summary>
            public virtual ProductItem Product { get; private set; }
            /// <summary>
            /// 订单编号
            /// </summary>
            public virtual string OrderNo { get; set; }
            /// <summary>
            /// 订单明细编号
            /// </summary>
            public virtual int OrderSID { get; set; }

            /// <summary>
            /// 客户
            /// </summary>
            public virtual string CustCode { get; set; }

            #region 订单表内容
            protected int _OrderMid = 0;
            /// <summary>
            /// 订单ID
            /// </summary>
            public virtual int OrderMID
            {
                get
                {
                    return _OrderMid;
                }
                set { _OrderMid = value; }
            }

            protected int _OrderNumb = 0;
            /// <summary>
            /// 订单总数
            /// </summary>
            public virtual int OrderNumb
            {
                get
                {
                    return _OrderNumb;
                }
                set { _OrderNumb = value; }
            }

            protected string _CustOrderNo = "";
            /// <summary>
            /// 客户单号
            /// </summary>
            public virtual string CustOrderNo
            {
                get
                {
                    return _CustOrderNo;
                }
                set { _CustOrderNo = value; }
            }

            protected string _OrderRemark = "";
            /// <summary>
            /// 订单备注
            /// </summary>
            public virtual string OrderRemark
            {
                get
                {
                    return _OrderRemark;
                }
                set { _OrderRemark = value; }
            }
            #endregion

            protected int _pNumb = 0;

            /// <summary>
            /// 生产数量
            /// </summary>
            public virtual int PNumb
            {
                get
                {
                    return _pNumb;
                }
                set
                {
                    _pNumb = value;
                }
            }

            /// <summary>
            /// 生产交期
            /// </summary>
            public virtual DateTime DeliveryDate { get; set; }
            /// <summary>
            /// 订单交期
            /// </summary>
            public virtual DateTime OrderDeliveryDate { get; set; }

            /// <summary>
            /// 输入时间
            /// </summary>
            public virtual DateTime InputDate { get; set; }

            /// <summary>
            /// 输入人
            /// </summary>
            public virtual string Inputer { get; set; }

            /// <summary>
            /// 审核日期
            /// </summary>
            public DateTime CheckDate { get; set; }

            /// <summary>
            /// 审核人
            /// </summary>
            public virtual string Checker { get; set; }

            /// <summary>
            /// 备注
            /// </summary>
            public virtual string Remark { get; set; }

            /// <summary>
            /// 补单次数，实时计算的。
            /// </summary>
            public virtual int ReDoTimes
            {
                get
                {
                    try
                    {
                        if (OriginalRdsNo != null && OriginalRdsNo.Length > 0 && OriginalRdsNo != RdsNo)
                        {
                            string SQL = "Select Count(*) as CountNumb from moProduce Where RdsNo=@RdsNo";
                            using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@RdsNo", OriginalRdsNo)))
                            {
                                MyData.MyDataRow mr = md.MyRows.FirstOrDefault();
                                return Convert.ToInt32(mr["CountNumb"]);
                            }
                        }
                        return 0;
                    }
                    catch (Exception ex)
                    {
                        return 0;
                    }
                }
            }

            /// <summary>
            /// 补单备注
            /// </summary>
            public virtual string RemarkReDo { get; set; }
            /// <summary>
            /// 改单备注
            /// </summary>
            public virtual string RemarkChange { get; set; }
            /// <summary>
            /// 核销备注
            /// </summary>
            public virtual string FinishRemark { get; set; }
            /// <summary>
            /// 核销人
            /// </summary>
            public virtual string FinishMan { get; set; }
            /// <summary>
            /// 刻交期备注
            /// </summary>
            public virtual string DeliveryRemark { get; set; }
            /// <summary>
            /// 刻交期时间
            /// </summary>
            public virtual DateTime DeliveryCheckDate { get; set; }
            /// <summary>
            /// 刻交期人
            /// </summary>
            public virtual string DeliveryChecker { get; set; }
            /// <summary>
            /// 原始工单号
            /// </summary>
            public virtual string OriginalRdsNo { get; set; }
            /// <summary>
            /// 不合格处理单
            /// </summary>
            public virtual string RejectRdsNo { get; set; }
            /// <summary>
            /// 退料单号
            /// </summary>
            public virtual string ReturnRdsNo { get; set; }
            /// <summary>
            /// 客诉单号
            /// </summary>
            public virtual string ComplainRdsNo { get; set; }
            /// <summary>
            /// 生产状态，-1：已砍单，0:未审核，1:已审核，2:生产中，3：生产完成
            /// </summary>
            public virtual int Status { get; set; }
            /// <summary>
            /// 生产状态
            /// </summary>
            public virtual string StatusWord
            {
                get
                {
                    if (Status < 0 && StockStatus == 2)
                    {
                        return MyConvert.ZHLC("已核销");
                    }
                    else
                    {
                        switch (Status)
                        {
                            case 0:
                               return MyConvert.ZHLC("未审核");
                            case 1:
                                return MyConvert.ZHLC("已审核");
                            case 2:
                                return MyConvert.ZHLC("生产中");
                            case 3:
                                return MyConvert.ZHLC("生产完成");
                            default:
                                return MyConvert.ZHLC("状态异常");
                        }
                    }
                }
            }

            /// <summary>
            /// 入库状态，0：未入库，1：入库中，2：入库完成
            /// </summary>
            public virtual int StockStatus { get; set; }
            /// <summary>
            /// 入库状态词语
            /// </summary>
            public virtual string StockStatusWord
            {
                get
                {
                    if (Status < 0 && StockStatus == 2)
                    {
                        return MyConvert.ZHLC("已核销");
                    }
                    else
                    {
                        switch (StockStatus)
                        {
                            case 0:
                                return MyConvert.ZHLC("未入库");
                            case 1:
                                return MyConvert.ZHLC("入库中");
                            case 2:
                                return MyConvert.ZHLC("入库完成");
                            default:
                                return MyConvert.ZHLC("状态异常");
                        }
                    }
                }
            }
            /// <summary>
            /// 核销以前的工单状态
            /// </summary>
            public int BeforeStatusID { get; set; }
            /// <summary>
            /// 完工日期
            /// </summary>
            public virtual DateTime FinishDate { get; set; }
            /// <summary>
            /// 入库完成日期
            /// </summary>
            public virtual DateTime StockFinishDate { get; set; }
            /// <summary>
            /// 工单性质ID
            /// 0:普通工单，1：补单，2：打样单，3：改单，4：重工/挑拣，5：拆解，6：性质异常
            /// </summary>
            public virtual int PropertyID { get; set; }
            /// <summary>
            /// 工单性质
            /// </summary>
            public virtual string Property
            {
                get
                {
                    switch (PropertyID)
                    {
                        case 0:
                            return MyConvert.ZHLC("普通工单");
                        case 1:
                            return MyConvert.ZHLC("补单");
                        case 2:
                            return MyConvert.ZHLC("打样单");
                        case 3:
                            return MyConvert.ZHLC("改单");
                        case 4:
                            return MyConvert.ZHLC("重工/挑拣");
                        case 5:
                            return MyConvert.ZHLC("拆解");
                        default:
                            return MyConvert.ZHLC("性质异常");
                    }
                }
            }
            /// <summary>
            /// 新旧版
            /// </summary>
            public virtual string VerType { get; set; }
            /// <summary>
            /// 最后工序
            /// </summary>
            public virtual string LastProcessCode { get; set; }

            /// <summary>
            /// 最后工序ID
            /// </summary>
            public virtual int LastProcessID { get; set; }

            /// <summary>
            /// 应用库存数
            /// </summary>
            public virtual float UseStockNumb { get; set; }

            /// <summary>
            /// 是否保税
            /// </summary>
            public virtual bool isBonded { get; set; }

            /// <summary>
            /// 入库库位类型
            /// </summary>
            public virtual string InStockLocationType { get; set; }

            #region 各种版号
            protected string _PlateCode;

            public virtual string PlateCode
            {
                get
                {
                    return (_PlateCode != null && _PlateCode.Length > 0) ? _PlateCode : (Product != null) ? Product.PlateCode : null;
                }
                set
                {
                    _PlateCode = value;
                }
            }


            protected string _CutPlateCode;

            public virtual string CutPlateCode
            {
                get
                {
                    return (_CutPlateCode != null && _CutPlateCode.Length > 0) ? _CutPlateCode : (Product != null) ? Product.CutPlateCode : null;
                }
                set
                {
                    _CutPlateCode = value;
                }
            }

            protected string _HotPlateCode;

            public virtual string HotPlateCode
            {
                get
                {
                    return (_HotPlateCode != null && _HotPlateCode.Length > 0) ? _HotPlateCode : (Product != null) ? Product.HotPlateCode : null;
                }
                set
                {
                    _HotPlateCode = value;
                }
            }

            protected string _ScreenPlateCode;

            public virtual string ScreenPlateCode
            {
                get
                {
                    return (_ScreenPlateCode != null && _ScreenPlateCode.Length > 0) ? _ScreenPlateCode : (Product != null) ? Product.ScreenPlateCode : null;
                }
                set
                {
                    _ScreenPlateCode = value;
                }
            }
            #endregion

            #endregion 属性存储器

            #region 集合属性
            /// <summary>
            /// 部件
            /// </summary>
            protected List<PartItem> _ProduceParts = new List<PartItem>();

            public virtual List<PartItem> Parts
            {
                get { return _ProduceParts; }
                set { _ProduceParts = value; }
            }
            /// <summary>
            /// 物料集合
            /// </summary>
            protected List<ProduceMaterial> _ProduceMaterials = new List<ProduceMaterial>();
            /// <summary>
            /// 物料集合
            /// </summary>
            public virtual List<ProduceMaterial> Materials
            {
                get { return _ProduceMaterials; }
                set { _ProduceMaterials = value; }
            }
            /// <summary>
            /// 工序集合
            /// </summary>
            protected List<ProduceProcess> _ProduceProcesses = new List<ProduceProcess>();
            /// <summary>
            /// 工序集合
            /// </summary>
            public virtual List<ProduceProcess> Processes
            {
                get { return _ProduceProcesses; }
                set { _ProduceProcesses = value; }
            }

            protected List<ProduceTogether> _ProduceTogether = new List<ProduceTogether>();

            /// <summary>
            /// 工单合单集合
            /// </summary>
            public virtual List<ProduceTogether> Toghethers
            {
                get { return _ProduceTogether; }
                set { _ProduceTogether = value; }
            }

            protected List<ProduceDestroyProduct> _ProduceDestroyProduct = new List<ProduceDestroyProduct>();

            /// <summary>
            /// 工单拆解入库子阶料件表
            /// </summary>
            public virtual List<ProduceDestroyProduct> DestroyProduct
            {
                get { return _ProduceDestroyProduct; }
                set { _ProduceDestroyProduct = value; }
            }

            #endregion

            /// <summary>
            /// 深拷贝
            /// </summary>
            /// <returns></returns>
            public ProduceNoteItem DeepClone()
            {
                return this.Clone();
                /*
                using (Stream objectStream = new MemoryStream())
                {
                    IFormatter formatter = new BinaryFormatter();
                    formatter.Serialize(objectStream, (ProduceNoteItem)this);
                    objectStream.Seek(0, SeekOrigin.Begin);
                    return formatter.Deserialize(objectStream) as ProduceNoteItem;
                }
                */
            }
        }

        /// <summary>
        /// 工单物料类
        /// </summary>
        [Serializable]
        public class ProduceMaterial
        {
            public ProduceMaterial(ProduceNote xParant)
            {
                Parent = xParant;
            }

            public ProduceMaterial(MyData.MyDataRow a, ProduceNote xParent)
            {
                ID = a.IntValue("id");
                zbid = a.IntValue("zbid");
                PartID = a.Value("PartID");
                PartNumb = a.IntValue("PartNumb") <= 0 ? xParent.PNumb : a.IntValue("PartNumb");
                TypeCode = a.Value("TypeCode");
                TypeName = a.Value("TypeName");
                Code = a.Value("Code");
                Name = a.Value("Name");
                Remark = a.Value("Remark");
                Length = a.Value<double>("Length");
                Width = a.Value<double>("Width");
                PlateNumber = a.IntValue("PlateNumb");
                CutNumb = a.IntValue("CutNumb");
                ColNumb = a.IntValue("ColNumb");
                ProcCode = a.Value("ProcCode");
                ProcID = a.IntValue("ProcID");
                PrintLossNumb = a.Value<decimal>("PrintLossNumb");
                PostpressLossNumb = a.Value<decimal>("PostpressLossNumb");
                SheetNumber = a.Value<decimal>("SheetNumb");
                Number = a.Value<double>("Number");
                OverflowNumber = a.Value<double>("OverflowNumb");
                PickedNumber = a.Value<double>("PickedNumb");
                RetrunNumber = a.Value<double>("ReturnNumb");
                UseStockNumb = a.IntValue("UseStockNumb");
                UnitNumb = a.Value<double>("sNumb");
                Specs = a.Value("Specs");
                PureName = a.Value("PureName");
                OriginalLength = a.Value<double>("OLength");
                OriginalWidth = a.Value<double>("OWidth");
                Parent = xParent;
                Unit = a.Value("unit");
                MaterialPickingType = (BomMaterial.MaterialPickingTypeEnum)a.IntValue("PickingType");
                MaterialType = GetMaterialBomType(a.IntValue("TP").ToString(), a.Value("TypeCode"), a.Value("Name"));
            }
            public ProduceNote Parent { get; set; }

            public BomMaterial.BomMaterialTypeEnum MaterialType { get; set; }
            /// <summary>
            /// 类型
            /// </summary>
            public string MaterialTypeWord
            {
                get
                {
                    string[] r = new string[] { "纸张", "半成品", "浪纸", "膜", "油墨", "PS版", "刀模", "物料" };
                    return r[(int)MaterialType];
                }
            }

            public int zbid { get; set; }
            public int ID { get; set; }
            /// <summary>
            /// 归属工序的ID
            /// </summary>
            public int ProcID { get; set; }
            /// <summary>
            /// 归属工序
            /// </summary>
            public string ProcCode { get; set; }
            /// <summary>
            /// 歸屬工序
            /// </summary>
            public ProduceProcess Process
            {
                get
                {
                    if (Parent != null && Parent.Processes != null)
                    {
                        var v = from a in Parent.Processes
                                where a.ID == ProcID
                                select a;
                        return v.FirstOrDefault();
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            /// <summary>
            /// 备注
            /// </summary>
            public string Remark { get; set; }
            /// <summary>
            /// 分类
            /// </summary>
            public string TypeCode { get; set; }
            /// <summary>
            /// 分类名称
            /// </summary>
            public string TypeName { get; set; }
            /// <summary>
            /// 部件名称
            /// </summary>
            public string PartID { get; set; }
            /// <summary>
            /// 物料编号
            /// </summary>
            public string Code { get; set; }
            /// <summary>
            /// 物料名称
            /// </summary>
            public string Name { get; set; }

            private double _numb;
            /// <summary>
            /// 领料数量。领料张数Ceiling((SheetNumber+PrintLossNumb+PostpressLossNumb)/CutNumb)
            /// </summary>
            public double Number
            {
                get
                {
                    if (Parent != null && (Parent.isEdit || _numb == 0))
                    {
                        if (MaterialType == BomMaterial.BomMaterialTypeEnum.BomWavePaper || MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper)
                        {
                            _numb = Math.Ceiling(Convert.ToDouble(SheetNumber + PrintLossNumb + PostpressLossNumb) / Convert.ToDouble(CutNumb));
                        }
                        else if (this.MaterialType == BomMaterial.BomMaterialTypeEnum.BomINK && this.Process != null)
                        {
                            _numb = Convert.ToDouble((Convert.ToDouble(UnitNumb) * (Convert.ToDouble(this.Process.PNumb) + Convert.ToDouble(this.Process.WastageNumb)) / 1000.00) * 1.1);
                            PartNumb = this.Process.PNumb + this.Process.WastageNumb;
                        }
                        else
                        {
                            if (Material.IsNotNull() && Material.NumberInteger)
                                _numb = Math.Ceiling(Math.Round(Convert.ToDouble(UnitNumb) * Convert.ToDouble(PartNumb + PostpressLossNumb), 1));
                            else
                                _numb = Convert.ToDouble(UnitNumb) * Convert.ToDouble(PartNumb + PostpressLossNumb);
                        }
                    }
                    if (this.MaterialType == BomMaterial.BomMaterialTypeEnum.BomINK)
                        return _numb * 1000.00;
                    else
                        return _numb;
                }
                set
                {
                    _numb = value;
                }
            }

            private decimal _sheetnumber = 0;
            /// <summary>
            /// 印纸张数/计算张数Ceiling（PartNumb/ColNumb)
            /// </summary>
            public decimal SheetNumber
            {
                get
                {
                    if (Parent != null && (Parent.isEdit || _sheetnumber == 0))
                    {
                        _sheetnumber = Convert.ToDecimal(Math.Ceiling(Convert.ToDouble(PartNumb - UseStockNumb) / Convert.ToDouble(ColNumb)));
                    }
                    return _sheetnumber;
                }
                set
                {
                    _sheetnumber = value;
                }
            }

            private int _cutNumb = 1;
            /// <summary>
            /// 开数
            /// </summary>
            public int CutNumb
            {
                get { return _cutNumb <= 0 ? 1 : _cutNumb; }
                set
                {
                    _cutNumb = value;
                    if (value <= 0) _cutNumb = 1;
                }
            }

            private int _colNumb = 1;
            /// <summary>
            /// 模数
            /// </summary>
            public int ColNumb
            {
                get { return _colNumb <= 0 ? 1 : _colNumb; }
                set
                {
                    _colNumb = value;
                    if (value <= 0) _colNumb = 1;
                }
            }

            int _APETColNumb = 0, _APETCutNumb = 0;

            public int APETCutNumb
            {
                get
                {
                    if (_APETCutNumb == 0) MakeAPETInfo();
                    return _APETCutNumb;
                }
            }

            public int APETColNumb
            {
                get
                {
                    if (_APETColNumb == 0) MakeAPETInfo();
                    return _APETColNumb;
                }
            }

            public void MakeAPETInfo()
            {
                string mword = Remark;
                Regex rgx = new Regex(@"\b(\d+?)[\/\\](\d+?)[\/\\](\d+?)\b");
                if (mword.Length > 0 && rgx.IsMatch(mword))
                {
                    rgx = new Regex(@"[\\\/]");
                    string[] xWords = rgx.Split(mword);
                    if (xWords.Length >= 3)
                    {
                        string cword = xWords[2], xword = xWords[1];
                        if (!int.TryParse(cword, out _APETColNumb)) _APETColNumb = 1;
                        if (!int.TryParse(xword, out _APETCutNumb)) _APETCutNumb = 1;
                    }
                }
            }


            /// <summary>
            /// 长
            /// </summary>
            public double Length { get; set; }
            /// <summary>
            /// 宽
            /// </summary>
            public double Width { get; set; }
            /// <summary>
            /// 尺寸
            /// </summary>
            public string Size { get { return string.Format("{0}x{1}", Length, Width); } }

            /// <summary>
            /// 原纸长
            /// </summary>
            public double OriginalLength { get; set; }
            /// <summary>
            /// 原纸宽
            /// </summary>
            public double OriginalWidth { get; set; }
            /// <summary>
            /// 版数
            /// </summary>
            public int PlateNumber { get; set; }
            /// <summary>
            /// 生产批量
            /// </summary>
            public int PartNumb { get; set; }

            /// <summary>
            /// 印刷放数,只有纸张有，单位是印纸张数
            /// </summary>
            public decimal PrintLossNumb { get; set; }
            /// <summary>
            /// 后制放数，单位是印纸张数
            /// </summary>
            public decimal PostpressLossNumb { get; set; }

            private double _UnitNumb = 0;
            /// <summary>
            /// 单位用量 字段sNumb,Number/PartNumb
            /// </summary>
            public double UnitNumb
            {
                get
                {
                    if (Parent != null && (Parent.isEdit || _UnitNumb == 0))
                    {
                        if (MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper || MaterialType == BomMaterial.BomMaterialTypeEnum.BomWavePaper)
                        {
                            if (_UnitNumb <= 0) _UnitNumb = Convert.ToDouble(Number) / Convert.ToDouble(PartNumb);
                        }
                        //if (_dosageNumb == 0) _dosageNumb = 1;
                    }
                    if (MaterialType == BomMaterial.BomMaterialTypeEnum.BomINK)
                        return _UnitNumb * 1000;
                    else
                        return _UnitNumb;
                }
                set
                {
                    _UnitNumb = value;
                }
            }
            /// <summary>
            /// 使用残页数,只有纸张有，而且不可以在纸张里边改。
            /// </summary>
            public int UseStockNumb { get; set; }

            public string Specs { get; set; }

            public string PureName { get; set; }

            /// <summary>
            /// 部件加工还是成品加工
            /// </summary>
            public PartItem.PartAssmeblyType PartType
            {
                get
                {
                    return PartID.Length > 0 ? PartItem.PartAssmeblyType.Part : PartItem.PartAssmeblyType.Assmebly;
                }
            }

            public PartItem Part
            {
                get
                {
                    if (Parent != null && Parent.Parts != null && this.PartID != null && this.PartID.Length > 0)
                    {
                        var v = from a in Parent.Parts
                                where a.PartID == this.PartID
                                select a;
                        if (v.Count() > 0) return v.FirstOrDefault();
                    }
                    return null;
                }
            }
            /// <summary>
            /// 是否绑定到一行
            /// </summary>
            public bool isBinding { get; set; }

            /// <summary>
            /// 纸张浪费率
            /// </summary>
            public double PaperLostPercent
            {
                get
                {
                    if (OriginalWidth > 0 && OriginalLength > 0)
                    {
                        double p = (Convert.ToDouble(Length) * Convert.ToDouble(Width) * Convert.ToDouble(CutNumb)) / (Convert.ToDouble(OriginalLength) * Convert.ToDouble(OriginalWidth));
                        return 1.000 - p;
                    }
                    else
                    {
                        return 0;
                    }
                }
            }
            /// <summary>
            /// 已领物料数量，多部件领料数大于本部件则认为已领。
            /// 这是库存单位数量，不是核算后的成品数。
            /// </summary>
            public double PickedNumber { get; set; }
            /// <summary>
            /// 退料数量
            /// </summary>
            public double RetrunNumber { get; set; }
            /// <summary>
            /// 增料数量
            /// </summary>
            public double OverflowNumber { get; set; }
            /// <summary>
            /// 单位
            /// </summary>
            public string Unit { get; set; }

            public IMaterialItem Material
            {
                get
                {
                    if (Materials.Contains(Code))
                        return Materials[Code];
                    else if (Products.Contains(Code))
                        return Products[Code];
                    else
                        return null;
                }
            }

            public string Sort_ID
            {
                get
                {
                    return string.Format("{0}_{1}_{2}", PartID.Length > 0 ? 100 : 200,
                        PartID.Length > 0 ? PartID : "Default",
                        Convert.ToString(ID).PadLeft(10, '0'));
                }
            }
            /// <summary>
            /// 领料方式
            /// </summary>
            public BomMaterial.MaterialPickingTypeEnum MaterialPickingType { get; set; }

        }
        /// <summary>
        /// 工单工艺表
        /// </summary>
        [Serializable]
        public class ProduceProcess
        {
            public ProduceProcess(ProduceNote xParant)
            {
                Parent = xParant;
            }
            public ProduceProcess(MyData.MyDataRow a, ProduceNote xParant)
            {
                ID = a.IntValue("id");
                zbid = a.IntValue("zbid");
                PartID = a.Value("PartID");
                Process = Processes[a.Value("Code")];
                Machine = Machines[a.Value("MachineCode")];
                fColorWord = a.Value("catxt");
                bColorWord = a.Value("cbtxt");
                ColNumb = a.IntValue("ColNumb");
                Length = a.Value<double>("Length");
                Width = a.Value<double>("Width");
                ProcMemo = a.Value("ProcMemo");
                OtherMemo = a.Value("OtherMemo");
                CleanMachineCoefficient = a.Value<double>("jN");
                ProductionCoefficient = a.Value<double>("pN");
                PNumb = a.IntValue("ReqNumb");
                UseRemainingNumb = a.IntValue("UseStockNumb");
                WaitTime = a.Value<double>("WaitTime");
                WastageNumb = a.IntValue("WastageNumb");
                FinishNumb = a.IntValue("FinishNumb");
                RejectNumb = a.IntValue("RejectNumb");
                LossedNumb = a.IntValue("LossedNumb");
                OverMan = a.Value("OverMan");
                OverRemark = a.Value("OverRemark");
                isOut = a.BooleanValue("isOut");
                isFinished = (a.DateTimeValue("OverDate") > DateTime.MinValue);
                FinishDate = a.DateTimeValue("OverDate");
                FirstPrintSide = a.IntValue("FirstPrintSide");
                FirstSideFinishNumb = a.Value<double>("FirstSideFinishNumb");
                AdjustNumb = a.IntValue("AdjustNumb");
                SampleNumb = a.IntValue("SampleNumb");
                Parent = xParant;
            }

            public ProcessItem Process { get; set; }

            public MachineItem Machine { get; set; }

            public ProduceNote Parent { get; set; }

            public int ID { get; set; }

            public int zbid { get; set; }

            public int Sort_Assembly { get; set; }
            /// <summary>
            /// 排序ID，平时不用，临时使用
            /// </summary>
            public int Sort_ID { get; set; }

            public string PartID { get; set; }

            private int _CutNumb = 1;
            /// <summary>
            /// 开数
            /// </summary>
            public int CutNumb { get { return _CutNumb <= 0 ? 1 : _CutNumb; } set { _CutNumb = value; } }
            private int _ColNumb = 1;
            /// <summary>
            /// 模数
            /// </summary>
            public int ColNumb { get { return _ColNumb <= 0 ? 1 : _ColNumb; } set { _ColNumb = value; } }
            /// <summary>
            /// 加工次数
            /// </summary>
            public int PrintTimes { get; set; }
            private double _ProductionCoefficient = 1;
            /// <summary>
            /// 生产难度系数
            /// </summary>
            public double ProductionCoefficient
            {
                get { return _ProductionCoefficient <= 0 ? 1 : _ProductionCoefficient; }
                set { _ProductionCoefficient = value; }
            }
            private double _CleanMachineCoefficient = 1;
            /// <summary>
            /// 洗车时间系数
            /// </summary>
            public double CleanMachineCoefficient
            {
                get { return _CleanMachineCoefficient <= 0 ? 1 : _CleanMachineCoefficient; }
                set { _CleanMachineCoefficient = value; }
            }
            /// <summary>
            /// 长
            /// </summary>
            public double Length { get; set; }
            /// <summary>
            /// 宽
            /// </summary>
            public double Width { get; set; }
            /// <summary>
            /// 尺寸
            /// </summary>
            public string Size { get { return string.Format("{0}x{1}", Length, Width); } }
            /// <summary>
            /// 工艺要求
            /// </summary>
            public string ProcMemo { get; set; }
            /// <summary>
            /// 生产备注
            /// </summary>
            public string OtherMemo { get; set; }
            /// <summary>
            /// 等干时间
            /// </summary>
            public double WaitTime { get; set; }

            /// <summary>
            /// 部件加工还是成品加工
            /// </summary>
            public PartItem.PartAssmeblyType PartType
            {
                get
                {
                    return PartID.Length > 0 ? PartItem.PartAssmeblyType.Part : PartItem.PartAssmeblyType.Assmebly;
                }
            }

            public PartItem Part
            {
                get
                {
                    if (Parent != null && Parent.Parts != null && this.PartID != null && this.PartID.Length > 0)
                    {
                        var v = from a in Parent.Parts
                                where a.PartID == this.PartID
                                select a;
                        if (v.Count() > 0) return v.FirstOrDefault();
                    }
                    return null;
                }
            }
            /// <summary>
            /// 版数
            /// </summary>
            public int PlateNumber { get; set; }

            /// <summary>
            /// 工序应产数
            /// </summary>
            public int PNumb { get; set; }

            /// <summary>
            /// 工序损耗数
            /// </summary>
            public int WastageNumb { get; set; }

            /// <summary>
            /// 预期加工时间
            /// </summary>
            public double RunTime
            {
                get
                {
                    try
                    {
                        return ((PNumb + WastageNumb) / (Machine.StdCapacity * this.ProductionCoefficient)) * 60 + (Machine.StdPerpareTime * CleanMachineCoefficient);
                    }
                    catch
                    {
                        return 0;
                    }
                }
            }

            /// <summary>
            /// 计算上工序
            /// </summary>
            public void CalcuatePreviousProcess()
            {
                if (PreviousProcess != null)
                {
                    double pNumb = Convert.ToDouble(((this.PNumb + this.WastageNumb) * this.ColNumb) - this.UseRemainingNumb);
                    var v = from a in Parent.Processes
                            where a.NextProcess == this
                            select a;
                    if (v.Count() > 0)
                    {
                        foreach (var item in v)
                        {
                            item.PNumb = Convert.ToInt32(Math.Ceiling(pNumb / (double)item.ColNumb));
                            item.CalcuatePreviousProcess();
                        }
                    }
                }
            }

            /// <summary>
            /// 应用残页数
            /// </summary>
            public int UseRemainingNumb { get; set; }
            /// <summary>
            /// 已经领料的残页数
            /// </summary>
            public int UseRemainingStockNumb { get; set; }
            /// <summary>
            /// 完工良品数，PCS产品数
            /// </summary>
            public int FinishNumb { get; set; }
            /// <summary>
            /// 完工不良品数,PCS
            /// </summary>
            public int RejectNumb { get; set; }
            /// <summary>
            /// 完工损耗数,样品数加过版纸数，PCS
            /// </summary>
            public int LossedNumb { get; set; }
            /// <summary>
            /// 过版纸数
            /// </summary>
            public int AdjustNumb { get; set; }
            /// <summary>
            /// 样品数
            /// </summary>
            public int SampleNumb { get; set; }
            /// <summary>
            /// 完工或者被核销
            /// </summary>
            public bool isFinished { get; set; }
            /// <summary>
            /// 核销人
            /// </summary>
            public string OverMan { get; set; }
            /// <summary>
            /// 核销原因
            /// </summary>
            public string OverRemark { get; set; }
            /// <summary>
            /// 产生的残页数
            /// </summary>
            public int RemainingNumb { get; set; }
            /// <summary>
            /// 已经入库的残页数
            /// </summary>
            public int RemainingStockNumb { get; set; }


            private double _picknumb;

            /// <summary>
            /// 上工序来料数，当首工序的时候是领料数。来料数和领料数都是PCS数。
            /// </summary>
            public double PickNumb
            {
                get
                {
                    if (Parent != null && Parent.Product.isKIT && CompanyType == "MT")
                    {
                        _picknumb = PNumb;
                    }
                    else
                    {
                        if (this.PreviousProcess != null)
                        {
                            #region 有上工序，非领料工序
                            if (this.PartID.Length <= 0)
                            {  //下工序等于本工序的最小完工数
                                var vpp = from a in Parent.Processes
                                          where a.ID < this.ID && a.NextProcess != null && a.NextProcess.ID == this.ID
                                          orderby a.ID descending
                                          select a.FinishNumb + a.UseRemainingNumb - a.RemainingNumb;
                                if (vpp.Count() > 0)
                                    _picknumb = vpp.Min();
                                else
                                    _picknumb = 0;
                            }
                            else
                            {
                                ProduceProcess PreProcess = this.PreviousProcess;
                                _picknumb = PreProcess.FinishNumb + PreProcess.UseRemainingNumb - PreProcess.RemainingNumb;
                            }
                            #endregion
                        }
                        else
                        {//起始工序，领料数
                            #region 领料工序
                            LoadPickMaterialInfo();
                            #endregion
                        }
                    }
                    return _picknumb;
                }
            }
            /// <summary>
            /// 加载所有领料的内容，和领料数量
            /// </summary>
            public void LoadPickMaterialInfo()
            {
                var vmm = from a in Parent.Materials
                          where a.ProcID == ID && a.PartID == PartID
                          select a;
                if (vmm != null && vmm.Count() > 0)
                {//有领料
                    #region 有领料，正常新工单
                    var vm1 = from a in vmm
                              where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper
                              select a;
                    if (vm1 != null && vm1.Count() > 0)
                    {//领料是纸张
                        var vmp1 = from a in vmm
                                   where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper
                                   select (a.PickedNumber + a.OverflowNumber - a.RetrunNumber) * a.CutNumb * a.ColNumb;
                        _picknumb = vmp1.Sum();
                        _PickMaterial = vm1.FirstOrDefault();
                    }
                    else
                    {
                        var vm2 = from a in vmm
                                  where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomWavePaper
                                  select a;
                        if (vm2 != null && vm2.Count() > 0)
                        {//领料是浪纸
                            var vmp2 = from a in vmm
                                       where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomWavePaper
                                       select (a.PickedNumber + a.OverflowNumber - a.RetrunNumber) * a.ColNumb;
                            _picknumb = vmp2.Sum();
                            _PickMaterial = vm2.FirstOrDefault();
                        }
                        else
                        {
                            var vm3 = from a in vmm
                                      where a.PureName.Contains("APET")
                                      select a;
                            if (vm3 != null && vm3.Count() > 0)
                            {//这是APET
                                var vmp3 = from a in vmm
                                           where a.PureName.Contains("APET")
                                           select a.Number * a.APETColNumb * a.APETCutNumb;
                                _picknumb = Convert.ToDouble(vmp3.Sum());
                                _PickMaterial = vm3.FirstOrDefault();
                            }
                            else
                            {//虽然领料，但是什么都不是,返回应产数
                                //_picknumb = (PNumb + WastageNumb) * ColNumb;
                                //_PickMaterial = vmm.FirstOrDefault();
                                if (this.Parent.InputDate < Convert.ToDateTime("2014-11-18").Date)
                                {
                                    _picknumb = (PNumb + WastageNumb) * ColNumb;
                                    _PickMaterial = vmm.FirstOrDefault();
                                }
                                else
                                {
                                    //if (MyServiceLogic.LoginAccBook == MySimplyBasic.AcctBookSet.SZMY)
                                    //{
                                    //    string[] s = new string[7] { "414-M-0318", "414-M-0034", "414-M-0302", "414-M-0381", "MY-1117-0005", "S058NB-1117-0001", "S058NA-1117-0001" };
                                    //    var vxm = from a in vmm
                                    //              where a.UnitNumb != 0 && !s.Contains(a.Code)
                                    //              orderby Math.Ceiling((a.PickedNumber + a.OverflowNumber - a.RetrunNumber) / a.UnitNumb) ascending
                                    //              select a;
                                    //    if (vxm != null && vxm.Count() > 0)
                                    //    {
                                    //        _PickMaterial = vxm.FirstOrDefault();
                                    //        _picknumb = Math.Round((_PickMaterial.PickedNumber + _PickMaterial.OverflowNumber - _PickMaterial.RetrunNumber) / _PickMaterial.UnitNumb);
                                    //    }
                                    //    else
                                    //    {
                                    //        _picknumb = (PNumb + WastageNumb) * ColNumb;
                                    //        _PickMaterial = vmm.FirstOrDefault();
                                    //    }
                                    //}
                                    //else
                                    //{
                                    var vxm = from a in vmm
                                              where a.UnitNumb != 0 && ((int)a.MaterialPickingType) < ((int)BomMaterial.MaterialPickingTypeEnum.Mass)
                                              orderby Math.Ceiling((a.PickedNumber + a.OverflowNumber - a.RetrunNumber) / a.UnitNumb) ascending
                                              select a;
                                    if (vxm != null && vxm.Count() > 0)
                                    {
                                        _PickMaterial = vxm.FirstOrDefault();
                                        _picknumb = Math.Round((_PickMaterial.PickedNumber + _PickMaterial.OverflowNumber - _PickMaterial.RetrunNumber) / _PickMaterial.UnitNumb);
                                    }
                                    else
                                    {
                                        _picknumb = (PNumb + WastageNumb) * ColNumb;
                                        _PickMaterial = vmm.FirstOrDefault();
                                    }
                                    //}
                                }
                            }
                        }
                    }
                    #endregion
                }
                else
                {//如果没有领料直接返回应产数
                    #region 找不到领料，旧工单。
                    var vm1 = from a in Parent.Materials
                              where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper && a.PartID == PartID
                              select a;
                    if (vm1 != null && vm1.Count() > 0)
                    {//领料是纸张
                        var vmp1 = from a in Parent.Materials
                                   where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper && a.PartID == PartID
                                   select (a.PickedNumber + a.OverflowNumber - a.RetrunNumber) * a.CutNumb * a.ColNumb;
                        _picknumb = vmp1.Sum();
                        _PickMaterial = vm1.FirstOrDefault();
                    }
                    else
                    {
                        var vm2 = from a in Parent.Materials
                                  where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomWavePaper && a.PartID == PartID
                                  select a;
                        if (vm2 != null && vm2.Count() > 0)
                        {//领料是浪纸
                            var vmp2 = from a in Parent.Materials
                                       where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomWavePaper && a.PartID == PartID
                                       select (a.PickedNumber + a.OverflowNumber - a.RetrunNumber) * a.ColNumb;
                            _picknumb = vmp2.Sum();
                            _PickMaterial = vm2.FirstOrDefault();
                        }
                        else
                        {
                            var vm3 = from a in Parent.Materials
                                      where a.PureName.Contains("APET")
                                      select a;
                            if (vm3 != null && vm3.Count() > 0)
                            {//这是APET
                                var vmp3 = from a in Parent.Materials
                                           where a.PureName.Contains("APET")
                                           select a.Number * a.APETColNumb * a.APETCutNumb;
                                _picknumb = Convert.ToDouble(vmp3.Sum());
                                _PickMaterial = vm3.FirstOrDefault();
                            }
                            else
                            {//虽然领料，但是什么都不是,返回应产数
                                if (ColNumb != 0)
                                    _picknumb = (PNumb + WastageNumb) * ColNumb;
                                else
                                    _picknumb = PNumb + WastageNumb;
                                _PickMaterial = vmm.FirstOrDefault();
                            }
                        }
                    }
                    #endregion
                }
            }

            private ProduceMaterial _PickMaterial;

            public ProduceMaterial PickMaterial
            {
                get
                {
                    if (_PickMaterial == null)
                    {
                        LoadPickMaterialInfo();
                    }
                    return _PickMaterial;
                }
            }


            string _fColorWord;
            /// <summary>
            /// 正面颜色
            /// </summary>
            public string fColorWord
            {
                get { return _fColorWord; }
                set
                {
                    _fColorWord = value;
                    MakeColorNumb();
                }
            }
            string _bColorWord;
            /// <summary>
            /// 反面颜色
            /// </summary>
            public string bColorWord
            {
                get { return _bColorWord; }
                set
                {
                    _bColorWord = value;
                    MakeColorNumb();
                }
            }

            private void MakeColorNumb()
            {
                int xa1 = 0, xa2 = 0, xb1 = 0, xb2 = 0;
                string[] gColors = new string[] { "C", "M", "Y", "K" };
                if (_fColorWord != null && _fColorWord.Length > 0)
                {
                    var va = _fColorWord.Split(',');
                    foreach (var sItem in va)
                    {
                        if (sItem.Length > 0)
                        {
                            if (gColors.Contains(sItem))
                                xa1++;
                            else
                                xa2++;
                        }
                    }
                }
                if (_bColorWord != null && _bColorWord.Length > 0)
                {
                    var vb = _bColorWord.Split(',');
                    foreach (var sItem in vb)
                    {
                        if (sItem.Length > 0)
                        {
                            if (gColors.Contains(sItem))
                                xb1++;
                            else
                                xb2++;
                        }
                    }
                }
                CA1 = xa1; CA2 = xa2; CB1 = xb1; CB2 = xb2;
            }
            /// <summary>
            /// 正面四色数
            /// </summary>
            public int CA1 { get; set; }
            /// <summary>
            /// 正面专色数
            /// </summary>
            public int CA2 { get; set; }
            /// <summary>
            /// 反面四色数
            /// </summary>
            public int CB1 { get; set; }
            /// <summary>
            /// 反面专色数
            /// </summary>
            public int CB2 { get; set; }

            public bool isOut { get; set; }
            /// <summary>
            /// 上工序
            /// </summary>
            public ProduceProcess PreviousProcess
            {
                get
                {
                    if (Parent != null && Parent.Product.isKIT)
                    {
                        return null;
                    }
                    else
                    {
                        if (this.Part != null)
                        {
                            var v = from a in Parent.Processes
                                    where (a.Part == null ||
                                           (a.Part != null && a.PartID != this.PartID && a.Part.PrintPartID == this.Part.PrintPartID && a.Process != this.Process) ||
                                           (a.PartID == this.PartID)
                                           ) && a.ID < this.ID
                                    orderby a.ID descending
                                    select a;
                            if (v.Count() > 0) return v.FirstOrDefault();
                        }
                        else
                        {
                            var v = from a in Parent.Processes
                                    where a.ID < this.ID
                                    orderby a.ID descending
                                    select a;
                            if (v.Count() > 0) return v.FirstOrDefault();
                        }
                    }
                    return null;
                }
            }
            /// <summary>
            /// 下工序
            /// </summary>
            public ProduceProcess NextProcess
            {
                get
                {
                    if (this.Part != null)
                    {
                        var v = from a in Parent.Processes
                                where (a.Part == null || a.PartID == this.PartID) && a.ID > this.ID
                                orderby a.ID ascending
                                select a;
                        if (v.Count() > 0) return v.FirstOrDefault();
                    }
                    else
                    {
                        var v = from a in Parent.Processes
                                where a.ID > this.ID
                                orderby a.ID ascending
                                select a;
                        if (v.Count() > 0) return v.FirstOrDefault();
                    }
                    return null;
                }
            }
            /// <summary>
            /// 是否绑定到一行
            /// </summary>
            public bool isBinding { get; set; }

            /// <summary>
            /// 完工时间
            /// </summary>
            public DateTime FinishDate { get; set; }

            public ProduceProcess DeepClone()
            {
                using (Stream objectStream = new MemoryStream())
                {
                    IFormatter formatter = new BinaryFormatter();
                    formatter.Serialize(objectStream, this);
                    objectStream.Seek(0, SeekOrigin.Begin);
                    return formatter.Deserialize(objectStream) as ProduceProcess;
                }
            }
            /// <summary>
            /// 先完成的那一面是哪一面。0正面，1反面
            /// </summary>
            public int FirstPrintSide { get; set; }
            /// <summary>
            /// 先完成的那一面，完工数是多少。
            /// </summary>
            public double FirstSideFinishNumb { get; set; }

        }
        /// <summary>
        /// 工单合版表
        /// </summary>
        [Serializable]
        public class ProduceTogether
        {
            public int ID { get; set; }
            public int zbid { get; set; }
            public string ProdNo { get; set; }
            public string OrderNo { get; set; }
            public int OrderPid { get; set; }
            public string Code { get; set; }
            public decimal pNumb { get; set; }
            public decimal Numb { get; set; }
            public string Remark { get; set; }
            public int ProdProcID { get; set; }
            public DateTime OrderDate { get; set; }
            public decimal OrderNumb { get; set; }
            public string OrderRemark { get; set; }
        }
        /// <summary>
        /// 工单拆解入库的子阶料件表
        /// </summary>
        [Serializable]
        public class ProduceDestroyProduct
        {
            public int ID { get; set; }
            public int zbid { get; set; }
            public string Code { get; set; }
            public string Name { get; set; }
            public double Numb { get; set; }
            public double UnitNumb { get; set; }
            public string Remark { get; set; }
        }


        #endregion 工单类

        #region BOM类

        #region BOM各种实体类
        /// <summary>
        /// BOM物料
        /// </summary>
        [Serializable]
        public class BomMaterial
        {
            [Serializable]
            public enum BomMaterialTypeEnum
            {
                /// <summary>
                /// 纸张
                /// </summary>
                BomPaper,
                /// <summary>
                /// 半成品
                /// </summary>
                BomProduct,
                /// <summary>
                /// 瓦楞纸
                /// </summary>
                BomWavePaper,
                /// <summary>
                /// 膜
                /// </summary>
                BomOPP,
                /// <summary>
                /// 油墨
                /// </summary>
                BomINK,
                /// <summary>
                /// 版
                /// </summary>
                BomPlate,
                /// <summary>
                /// 模切板
                /// </summary>
                BomCutPlate,
                /// <summary>
                /// 其他辅料
                /// </summary>
                BomOtherMaterial
            };

            public enum MaterialPickingTypeEnum
            {
                /// <summary>
                /// 排程领料
                /// </summary>
                Plan = 0,
                /// <summary>
                /// 工单领料
                /// </summary>
                ProduceNote = 1,
                /// <summary>
                /// 批量领料不影响完工单
                /// </summary>
                Mass = 2,
                /// <summary>
                /// 其他领料
                /// </summary>
                Other = 3
            }


            public BomMaterial()
            {

            }
            public BomMaterial(MyData.MyDataRow a)
            {
                Name = Convert.ToString(a["Name"]);
                Code = Convert.ToString(a["Code"]);
                TypeCode = Convert.ToString(a["TypeCode"]);
                TypeName = Convert.ToString(a["TypeName"]);
                Number = a["Numb"].ConvertTo<double>(0, true);
                ColNumb = a["ColNumb"].ConvertTo<int>(0, true);
                Length = a["Length"].ConvertTo<Single>(0, true);
                Width = a["Width"].ConvertTo<Single>(0, true);
                Remark = Convert.ToString(a["remark"]);
                PartID = Convert.ToString(a["PartID"]);
                ProcCode = Convert.ToString(a["ProcCode"]);
                ProcID = a["ProcID"].ConvertTo<int>(0, true);
                CutNumb = a["CutNumb"].ConvertTo<int>(0, true);
                OriginalLength = a["OriginalLength"].ConvertTo<Single>(0, true);
                OriginalWidth = a["OriginalWidth"].ConvertTo<Single>(0, true);
                ID = a["id"].ConvertTo<int>(0, true);
                _id = a["_id"].ConvertTo<int>(0, true);
                zbid = a["zbid"].ConvertTo<int>(0, true);
                PlateNumber = a["PlateNumb"].ConvertTo<int>(0, true);
                Unit = Convert.ToString(a["Unit"]);
                PureName = Convert.ToString(a["PureName"]);
                isBonded = Convert.ToBoolean(a["Bonded"]);
                MaterialPickingType = (MaterialPickingTypeEnum)a.IntValue("PickingType");
                MaterialType = GetMaterialBomType(Convert.ToString(a["TP"]), Convert.ToString(a["TypeCode"]), Convert.ToString(a["Name"]));
            }

            public Bom Parent { get; set; }

            public BomMaterialTypeEnum MaterialType { get; set; }

            public string MaterialTypeWord
            {
                get
                {
                    string[] r = new string[] { "纸张", "半成品", "浪纸", "膜", "油墨", "PS版", "刀模", "物料" };
                    return r[(int)MaterialType];
                }
            }

            public int _id { get; set; }
            public int zbid { get; set; }
            public int ID { get; set; }
            /// <summary>
            /// 所属工序ID
            /// </summary>
            public int ProcID { get; set; }
            public string ProcCode { get; set; }
            public string Remark { get; set; }

            public string TypeCode { get; set; }
            public string TypeName { get; set; }
            /// <summary>
            /// 自动生成的排序ID
            /// </summary>
            public string Sort_ID
            {
                get
                {
                    return string.Format("{0}_{1}_{2}", PartID.Length > 0 ? 100 : 200,
                        PartID.Length > 0 ? PartID : "Default",
                        Convert.ToString(ID).PadLeft(10, '0'));
                }
            }

            public string PartID { get; set; }
            /// <summary>
            /// 物料编号
            /// </summary>
            public string Code { get; set; }
            /// <summary>
            /// 物料名称，长名称含有尺寸规格
            /// </summary>
            public string Name { get; set; }
            /// <summary>
            /// 物料名称，短名称。
            /// </summary>
            public string PureName { get; set; }

            /// <summary>
            /// 配比
            /// </summary>
            public double Number { get; set; }

            /// <summary>
            /// 单位用量
            /// </summary>
            public double UnitNumber
            {
                get
                {
                    if (MaterialType == BomMaterialTypeEnum.BomPaper)
                    {
                        return 1.00 / (Convert.ToDouble(ColNumb) * Convert.ToDouble(CutNumb));
                    }
                    else if (MaterialType == BomMaterialTypeEnum.BomWavePaper)
                    {
                        return 1.00 / Convert.ToDouble(ColNumb);
                    }
                    else if (MaterialType == BomMaterialTypeEnum.BomINK)
                    {
                        return Number * 1000.00;
                    }
                    else
                    {
                        return Number;
                    }
                }
            }

            /// <summary>
            /// 开数
            /// </summary>
            public int CutNumb { get; set; }
            /// <summary>
            /// 模数
            /// </summary>
            public int ColNumb { get; set; }
            /// <summary>
            /// 长
            /// </summary>
            public float Length { get; set; }
            /// <summary>
            /// 宽
            /// </summary>
            public float Width { get; set; }
            /// <summary>
            /// 原纸长
            /// </summary>
            public float OriginalLength { get; set; }
            /// <summary>
            ///
            /// </summary>
            public float OriginalWidth { get; set; }

            /// <summary>
            /// 尺寸
            /// </summary>
            public string Size { get { return string.Format("{0}x{1}", Length, Width); } }
            /// <summary>
            /// 版数
            /// </summary>
            public int PlateNumber { get; set; }
            /// <summary>
            /// 纸张浪费率
            /// </summary>
            public double PaperLostPercent
            {
                get
                {
                    if (OriginalWidth > 0 && OriginalLength > 0)
                    {
                        double p = (Convert.ToDouble(Length) * Convert.ToDouble(Width) * Convert.ToDouble(CutNumb)) / (Convert.ToDouble(OriginalLength) * Convert.ToDouble(OriginalWidth));
                        return 1.000 - p;
                    }
                    else
                    {
                        return 0;
                    }
                }
            }

            /// <summary>
            /// 部件加工还是成品加工
            /// </summary>
            public PartItem.PartAssmeblyType PartType
            {
                get
                {
                    return PartID.Length > 0 ? PartItem.PartAssmeblyType.Part : PartItem.PartAssmeblyType.Assmebly;
                }
            }

            public PartItem Part
            {
                get
                {
                    if (Parent != null && Parent.Parts != null && this.PartID != null && this.PartID.Length > 0)
                    {
                        var v = from a in Parent.Parts
                                where a.PartID == this.PartID
                                select a;
                        if (v.Count() > 0) return v.FirstOrDefault();
                    }
                    return null;
                }
            }

            /// <summary>
            /// 深拷贝
            /// </summary>
            /// <returns></returns>
            public BomMaterial DeepClone()
            {
                using (Stream objectStream = new MemoryStream())
                {
                    IFormatter formatter = new BinaryFormatter();
                    formatter.Serialize(objectStream, this);
                    objectStream.Seek(0, SeekOrigin.Begin);
                    return formatter.Deserialize(objectStream) as BomMaterial;
                }
            }

            public bool isBinding { get; set; }

            public string Unit { get; set; }
            /// <summary>
            /// 保税
            /// </summary>
            public bool isBonded { get; set; }

            public MaterialPickingTypeEnum MaterialPickingType { get; set; }

        }
        /// <summary>
        /// ＢＯＭ工艺
        /// </summary>
        [Serializable]
        public class BomProcess
        {
            public BomProcess()
            {

            }
            public BomProcess(MyData.MyDataRow a)
            {
                CleanMachineCoefficient = a.Value<double>("jN"); // Convert.ToSingle(a["jN"]);
                ProductionCoefficient = a.Value<double>("pN"); // Convert.ToDouble(a["pN"]);
                string ProcessCode = a.Value("Code"); // Convert.ToString(a["Code"]);
                Process = Processes[ProcessCode];
                string MachineCode = a.Value("MachineCode"); // Convert.ToString(a["MachineCode"]);
                Machine = Machines[MachineCode];
                fColorWord = a.Value("catxt"); // Convert.ToString(a["catxt"]);
                bColorWord = a.Value("cbtxt"); // Convert.ToString(a["cbtxt"]);
                ColNumb = a.IntValue("ColNumb"); // Convert.ToInt32(a["ColNumb"]);
                Length = a.Value<double>("Length"); // Convert.ToSingle(a["Length"]);
                Width = a.Value<double>("Width"); // Convert.ToSingle(a["Width"]);
                ProcMemo = a.Value("ProcMemo"); // Convert.ToString(a["ProcMemo"]);
                OtherMemo = a.Value("OtherMemo"); // Convert.ToString(a["OtherMemo"]);
                PartID = a.Value("PartID"); // Convert.ToString(a["PartID"]);
                ID = a.IntValue("ID"); // Convert.ToInt32(a["ID"]);
                zbid = a.IntValue("zbid"); // Convert.ToInt32(a["zbid"]);
                WaitTime = a.Value<double>("WaitTime"); // Convert.ToDouble(a["WaitTime"]);
                PersonNumb = a.IntValue("PersonNumb"); // Convert.ToInt32(a["PersonNumb"]);
            }

            public ProcessItem Process { get; set; }

            public MachineItem Machine { get; set; }

            public Bom Parent { get; set; }

            public int ID { get; set; }

            public int zbid { get; set; }

            public int Sort_Assembly { get; set; }
            /// <summary>
            /// 排序ID，平时没有用，排序临时用
            /// </summary>
            public string Sort_ID
            {
                get
                {
                    return string.Format("{0}_{1}_{2}", PartID.Length > 0 ? 100 : 200,
                        PartID.Length > 0 ? PartID : "default",
                         Convert.ToString(ID).PadLeft(10, '0'));
                }
            }

            public string PartID { get; set; }
            /// <summary>
            /// 模数
            /// </summary>
            public int ColNumb { get; set; }
            /// <summary>
            /// 加工次数
            /// </summary>
            public int PrintTimes { get; set; }
            /// <summary>
            /// 生产难度系数
            /// </summary>
            public double ProductionCoefficient { get; set; }
            /// <summary>
            /// 洗车时间系数
            /// </summary>
            public double CleanMachineCoefficient { get; set; }
            /// <summary>
            /// 长
            /// </summary>
            public double Length { get; set; }
            /// <summary>
            /// 宽
            /// </summary>
            public double Width { get; set; }
            /// <summary>
            /// 尺寸
            /// </summary>
            public string Size { get { return string.Format("{0}x{1}", Length, Width); } }
            /// <summary>
            /// 工艺说明
            /// </summary>
            public string ProcMemo { get; set; }
            /// <summary>
            /// 其他说明
            /// </summary>
            public string OtherMemo { get; set; }
            /// <summary>
            /// 等干时间
            /// </summary>
            public double WaitTime { get; set; }
            /// <summary>
            /// 部件加工还是成品加工
            /// </summary>
            public PartItem.PartAssmeblyType PartType
            {
                get
                {
                    return PartID.Length > 0 ? PartItem.PartAssmeblyType.Part : PartItem.PartAssmeblyType.Assmebly;
                }
            }

            string _fColorWord;
            /// <summary>
            /// 正面颜色
            /// </summary>
            public string fColorWord
            {
                get { return _fColorWord; }
                set
                {
                    _fColorWord = value;
                    MakeColorNumb();
                }
            }
            string _bColorWord;
            /// <summary>
            /// 反面颜色
            /// </summary>
            public string bColorWord
            {
                get { return _bColorWord; }
                set
                {
                    _bColorWord = value;
                    MakeColorNumb();
                }
            }

            private void MakeColorNumb()
            {
                int xa1 = 0, xa2 = 0, xb1 = 0, xb2 = 0;
                string[] gColors = new string[] { "C", "M", "Y", "K" };
                if (_fColorWord != null && _fColorWord.Length > 0)
                {
                    var va = _fColorWord.Split(',');
                    foreach (var sItem in va)
                    {
                        if (sItem.Length > 0)
                        {
                            if (gColors.Contains(sItem))
                                xa1++;
                            else
                                xa2++;
                        }
                    }
                }
                if (_bColorWord != null && _bColorWord.Length > 0)
                {
                    var vb = _bColorWord.Split(',');
                    foreach (var sItem in vb)
                    {
                        if (sItem.Length > 0)
                        {
                            if (gColors.Contains(sItem))
                                xb1++;
                            else
                                xb2++;
                        }
                    }
                }
                CA1 = xa1; CA2 = xa2; CB1 = xb1; CB2 = xb2;
            }
            /// <summary>
            /// 正面四色数
            /// </summary>
            public int CA1 { get; set; }
            /// <summary>
            /// 正面专色数
            /// </summary>
            public int CA2 { get; set; }
            /// <summary>
            /// 反面四色数
            /// </summary>
            public int CB1 { get; set; }
            /// <summary>
            /// 反面专色数
            /// </summary>
            public int CB2 { get; set; }

            public PartItem Part
            {
                get
                {
                    if (Parent != null && Parent.Parts != null && this.PartID != null && this.PartID.Length > 0)
                    {
                        var v = from a in Parent.Parts
                                where a.PartID == this.PartID
                                select a;
                        if (v.Count() > 0) return v.FirstOrDefault();
                    }
                    return null;
                }
            }
            /// <summary>
            /// 版数
            /// </summary>
            public int PlateNumber { get; set; }

            /// <summary>
            /// 开数
            /// </summary>
            public int PaperCutNumb
            {
                get
                {
                    if (Parent != null && Parent.Materials != null)
                    {
                        List<BomMaterial> xList = Parent.Materials;
                        int pCutNumb = 0;
                        if (this.PartID.Length > 0)
                        {
                            var v = from a in xList
                                    where a.Part != null && a.Part.PrintPartID == this.Part.PrintPartID && a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper
                                    orderby a.ID
                                    select a.CutNumb;
                            if (v.Count() > 0) pCutNumb = v.FirstOrDefault();
                        }
                        else
                        {
                            var v = from a in xList
                                    where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper
                                    orderby a.ID
                                    select a.CutNumb;
                            if (v.Count() > 0) pCutNumb = v.FirstOrDefault();
                        }
                        return pCutNumb;
                    }
                    else
                    {
                        return 1;
                    }
                }
            }

            public int PaperColNumb
            {
                get
                {
                    if (Parent != null && Parent.Materials != null)
                    {
                        List<BomMaterial> xList = Parent.Materials;
                        int pColNumb = 0;
                        if (this.PartID.Length > 0)
                        {
                            var v = from a in xList
                                    where a.Part != null && a.Part.PrintPartID == this.Part.PrintPartID && a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper
                                    orderby a.ID
                                    select a.ColNumb;
                            if (v.Count() > 0) pColNumb = v.FirstOrDefault();
                        }
                        else
                        {
                            var v = from a in xList
                                    where a.MaterialType == BomMaterial.BomMaterialTypeEnum.BomPaper
                                    orderby a.ID
                                    select a.ColNumb;
                            if (v.Count() > 0) pColNumb = v.FirstOrDefault();
                        }
                        return pColNumb;
                    }

                    else
                    {
                        return 1;
                    }
                }
            }
            /// <summary>
            /// 上工序
            /// </summary>
            public BomProcess PreviousProcess
            {
                get
                {
                    if (Parent == null && this.Process.Code == "9050")
                    {
                        return null;
                    }
                    else
                    {
                        if (this.Part != null)
                        {
                            var v = from a in Parent.Processes
                                    where (a.Part == null ||
                                           (a.Part != null && a.PartID != this.PartID && a.Part.PrintPartID == this.Part.PrintPartID && a.Process != this.Process) ||
                                           (a.PartID == this.PartID)
                                           ) && a.ID < this.ID
                                    orderby a.ID descending
                                    select a;
                            if (v.Count() > 0) return v.FirstOrDefault();
                        }
                        else
                        {
                            var v = from a in Parent.Processes
                                    where a.ID < this.ID
                                    orderby a.ID descending
                                    select a;
                            if (v.Count() > 0) return v.FirstOrDefault();
                        }
                    }
                    return null;
                }
            }
            /// <summary>
            /// 下工序
            /// </summary>
            public BomProcess NextProcess
            {
                get
                {
                    if (this.Part != null)
                    {
                        var v = from a in Parent.Processes
                                where (a.Part == null || a.PartID == this.PartID) && a.ID > this.ID
                                orderby a.ID ascending
                                select a;
                        if (v.Count() > 0) return v.FirstOrDefault();
                    }
                    else
                    {
                        var v = from a in Parent.Processes
                                where a.ID > this.ID
                                orderby a.ID ascending
                                select a;
                        if (v.Count() > 0) return v.FirstOrDefault();
                    }
                    return null;
                }
            }


            /// <summary>
            /// 深拷贝
            /// </summary>
            /// <returns></returns>
            public BomProcess DeepClone()
            {
                using (Stream objectStream = new MemoryStream())
                {
                    IFormatter formatter = new BinaryFormatter();
                    formatter.Serialize(objectStream, this);
                    objectStream.Seek(0, SeekOrigin.Begin);
                    return formatter.Deserialize(objectStream) as BomProcess;
                }
            }
            //是否有绑定到表格一行。
            public bool isBinding { get; set; }

            /// <summary>
            /// 人数
            /// </summary>
            public int PersonNumb { get; set; }
        }

        #endregion

        #region BOM本体
        /// <summary>
        /// BOM本体
        /// </summary>
        [Serializable]
        public class Bom
        {
            #region 全部BOM的内容

            /// <summary>
            /// 所有BOM物料存储器
            /// </summary>
            private static List<BomMaterial> _BomMaterialList = new List<BomMaterial>();

            public static int BomMaterialListCount
            {
                get
                {
                    if (_BomMaterialList == null)
                        return 0;
                    else
                        return _BomMaterialList.Count();
                }
            }

            public static void ClearBomDetialList()
            {
                _BomMaterialList = null;
                _BomProcessList = null;
            }

            /// <summary>
            /// 所有BOM的物料，包含纸张瓦楞纸和物料。
            /// </summary>
            public static List<BomMaterial> BomMaterialList
            {
                get
                {
                    if (_BomMaterialList == null) _BomMaterialList = new List<BomMaterial>();
                    if (_BomMaterialList.Count() <= 0)
                    {
                        MyRecord.Say("请稍等，加载BOM物料，从数据库获取所有BOM物料。请等待");
                        try
                        {
                            _BomMaterialList = GetBomMaterialList();
                        }
                        finally
                        {
                        }
                    }
                    return _BomMaterialList;
                }
                set
                {
                    _BomMaterialList = value;
                }
            }

            /// <summary>
            ///
            /// </summary>
            /// <returns></returns>
            public static List<BomMaterial> GetBomMaterialList()
            {
                return GetBomMaterialList(-1);
            }

            public static List<BomMaterial> GetBomMaterialList(int BomMainID)
            {
                if (BomMainID != 0)
                {
                    string SQL = @" Exec [_PMC_BomSelectMaterial] @BomID ";
                    MyData.MyDataParameter mps = new MyData.MyDataParameter("@BomID", BomMainID <= 0 ? 0 : BomMainID, MyData.MyDataParameter.MyDataType.Int);
                    try
                    {
                        using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, mps))
                        {
                            if (md != null && md.MyRows.Count > 0)
                            {
                                if (md.MyRows.Count > 100)
                                {
                                    int iItemCount = 1, ItemCount = md.MyRows.Count;
                                    List<BomMaterial> r = new List<BomMaterial>();
                                    foreach (var item in md.MyRows)
                                    {
                                        r.Add(new BomMaterial(item));
                                        iItemCount++;
                                    }
                                    return r;
                                }
                                else
                                {
                                    var v = from a in md.MyRows
                                            select new BomMaterial(a);
                                    return v.ToList();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                return null;

            }

            public static int BomProcessListCount
            {
                get
                {
                    if (_BomProcessList == null)
                        return 0;
                    else
                        return _BomProcessList.Count();
                }
            }
            /// <summary>
            /// 所有BOM工艺存储器
            /// </summary>
            private static List<BomProcess> _BomProcessList = new List<BomProcess>();
            /// <summary>
            /// 所有BOM工艺，包含KIT和非KIT
            /// </summary>
            public static List<BomProcess> BomProcessList
            {
                get
                {
                    if (_BomProcessList == null) _BomProcessList = new List<BomProcess>();
                    if (_BomProcessList.Count() <= 0)
                    {
                        bool wpclose = false;
                        MyRecord.Say("请稍等，初始化数据。正在读取所有BOM的工序到本地。请等待");
                        try
                        {
                            _BomProcessList = GetBomProcessList();
                        }
                        finally
                        {
                        }
                    }
                    return _BomProcessList;
                }
                set
                {
                    _BomProcessList = value;
                }
            }

            public static List<BomProcess> GetBomProcessList()
            {
                return GetBomProcessList(-1);
            }

            public static List<BomProcess> GetBomProcessList(int BomMainID)
            {
                if (BomMainID != 0)
                {
                    string SQL = @"Exec _PMC_BomSelectProcess  @BomID ";
                    MyData.MyDataParameter mps = new MyData.MyDataParameter("@BomID", BomMainID <= 0 ? 0 : BomMainID, MyData.MyDataParameter.MyDataType.Int);
                    try
                    {
                        using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, mps))
                        {
                            if (md != null && md.MyRows.Count > 0)
                            {
                                if (md.MyRows.Count > 100)
                                {
                                    int iItemCount = 1, ItemCount = md.MyRows.Count;
                                    List<BomProcess> r = new List<BomProcess>();
                                    foreach (var item in md.MyRows)
                                    {
                                        r.Add(new BomProcess(item));
                                        iItemCount++;
                                    }
                                    return r;
                                }
                                else
                                {
                                    var v = from a in md.MyRows
                                            select new BomProcess(a);
                                    return v.ToList();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                return null;
            }

            #endregion

            #region 属性
            /// <summary>
            /// BomID
            /// </summary>
            public int ID { get; private set; }
            /// <summary>
            /// 产品编号
            /// </summary>
            public string Code { get; private set; }
            /// <summary>
            /// 备注
            /// </summary>
            public string Remark { get; set; }
            /// <summary>
            /// 是否多部件
            /// </summary>
            public bool MutiplePart { get; set; }
            /// <summary>
            /// 刀模编号
            /// </summary>
            public string CutPlateCode { get; set; }
            /// <summary>
            /// 网版号
            /// </summary>
            public string ScreenPlateCode { get; set; }
            /// <summary>
            /// 烫金版号
            /// </summary>
            public string HotPlateCode { get; set; }
            /// <summary>
            /// 版号
            /// </summary>
            public string PlateCode { get; set; }
            /// <summary>
            /// 制造分类
            /// </summary>
            public string mType { get; set; }
            /// <summary>
            /// 是否是新BOM。
            /// </summary>
            public bool isNewBom { get; set; }
            /// <summary>
            /// 是否在修改状态中
            /// </summary>
            public bool isOnEdit { get; set; }
            /// <summary>
            /// 状态
            /// </summary>
            public int Status { get; set; }

            public bool isNewERPBom { get; set; }

            public double Price2 { get; set; }

            public ProductItem Parent { get; set; }

            #endregion

            #region 构造函数

            public Bom(string Code)
            {
                this.Code = Code;
                isOnEdit = false;
                Parent = Products[Code];
                LoadBom(Code);
            }

            public Bom(ProductItem prod)
            {
                Parent = prod;
                LoadBom(prod);
            }

            private void LoadBom(string code)
            {
                string SQL = @"Select * from pbBom Where code=@code";
                using (MyData.MyDataTable md = new MyData.MyDataTable(SQL, new MyData.MyDataParameter("@code", this.Code)))
                {
                    if (md != null && md.MyRows.Count > 0)
                    {
                        MyData.MyDataRow r = md.MyRows.FirstOrOnlyRow;
                        ID = r.IntValue("ID");
                        Remark = r.Value("remark");
                        CutPlateCode = r.Value("dm");
                        MutiplePart = r.BooleanValue("mPart");
                        Status = r.IntValue("Status");
                        isNewERPBom = r.StringValue("FlagCode").Trim().IsNotEmptyOrWhiteSpace();
                        isNewBom = false;
                        LoadBom(ID);
                    }
                    else
                        isNewBom = true;
                }
            }

            public void LoadBom(ProductItem prod)
            {
                MyRecord.Say("从印件资料加载BOM基本信息");
                Code = prod.Code;
                isOnEdit = false;
                if (prod.BomID > 0)
                {
                    Remark = prod.BomRemark;
                    CutPlateCode = prod.CutPlateCode;
                    ScreenPlateCode = prod.ScreenPlateCode;
                    HotPlateCode = prod.HotPlateCode;
                    MutiplePart = prod.MutiplePart;
                    Status = prod.BomStatus;
                    PlateCode = prod.PlateCode;
                    mType = prod.mType;
                    isNewERPBom = prod.isNewERPBom;
                    this.ID = prod.BomID;
                    isNewBom = false;
                }
                else
                    isNewBom = true;
            }

            public void reLoadDetial()
            {
                MyRecord.Say(string.Format("{0},读取BOM详细信息。",Code));
                if (Parent.IsNull())
                {
                    LoadBom(this.Code);
                }
                else
                {
                    LoadBom(Parent);
                    LoadBom(this.ID);
                }
            }

            private void LoadBom(int id)
            {
                if (id > 0)
                {
                    if (!isOnEdit)
                    {
                        ID = id;
                        isNewBom = false;
                        MyRecord.Say("读取BOM工艺和物料信息");
                        #region 加载物料
                        int xBomMaterialListCount = BomMaterialListCount;
                        if (xBomMaterialListCount > 0)
                        {
                            var vM = from a in BomMaterialList
                                     where a.zbid == id
                                     orderby a._id ascending
                                     select a;
                            _Materials = vM.ToList();
                        }
                        else
                        {
                            MyRecord.Say("读取BOM工艺。");
                            List<BomMaterial> xSource = GetBomMaterialList(this.ID);
                            if (xSource.IsNotEmptySet())
                            {
                                var vM = from a in xSource
                                         where a.zbid == id
                                         orderby a._id ascending
                                         select a;
                                _Materials = vM.ToList();
                            }
                            else
                            {
                                _Materials = new List<BomMaterial>();
                            }
                        }
                        ///设置长辈
                        if (_Materials != null && _Materials.Count > 0) _Materials.ForEach(a => a.Parent = this);
                        #endregion


                        #region 加载工序
                        int xBomProcessListCount = BomProcessListCount;
                        if (xBomProcessListCount > 0)
                        {
                            var vP = from a in BomProcessList
                                     where a.zbid == id
                                     orderby a.PartID, a.ID
                                     select a;
                            _Processes = vP.ToList();
                        }
                        else
                        {
                            MyRecord.Say("读取BOM物料。");
                            List<BomProcess> xSource = GetBomProcessList(this.ID);
                            if (xSource.IsNotEmptySet())
                            {
                                var vP = from a in xSource
                                         where a.zbid == id
                                         orderby a.PartID, a.ID
                                         select a;
                                _Processes = vP.ToList();
                            }
                            else
                            {
                                _Processes = new List<BomProcess>();
                            }
                        }
                        ///设置长辈
                        if (_Processes != null && _Processes.Count > 0) _Processes.ForEach(a => a.Parent = this);
                        #endregion


                        #region 单独加载

                        #endregion


                        #region 集体加载 太慢暂时不用
                        //var vP = from a in BomProcessList
                        //         where a.zbid == id
                        //         orderby a.ID ascending
                        //         select a;
                        //this.Processes = vP.ToList();
                        //if (Processes != null) Processes.ForEach(new Action<BomProcess>(x => x.Parent = this));

                        //var vM = from a in BomMaterialList
                        //         where a.zbid == id
                        //         orderby a.ID ascending
                        //         select a;
                        //this.Materials = vM.ToList();
                        //if (Materials != null) Materials.ForEach(new Action<BomMaterial>(x => x.Parent = this));
                        #endregion
                        MyRecord.Say("生成BOM。");
                        LoadBomParts();
                    }
                }
                else
                {
                    isNewBom = true;
                }
            }

            public void LoadBomParts()
            {
                if (_Materials != null && _Processes != null)
                {
                    var vPart = from a in
                                    (from xa in _Processes select xa.PartID).Union(from xb in _Materials select xb.PartID)
                                where a.Length > 0
                                group a by a into g
                                select new PartItem(g.Key);
                    _Parts = vPart.ToList();
                }
                else if (_Processes != null)
                {
                    var vPart = from a in
                                    (from xa in _Processes select xa.PartID)
                                where a.Length > 0
                                group a by a into g
                                select new PartItem(g.Key);
                    _Parts = vPart.ToList();
                }
                else if (_Materials != null)
                {
                    var vPart = from a in
                                    (from xb in _Materials select xb.PartID)
                                where a.Length > 0
                                group a by a into g
                                select new PartItem(g.Key);
                    _Parts = vPart.ToList();
                }
                if (_Parts != null) MutiplePart = (_Parts.Count > 1);

                if (!MutiplePart)
                {
                    if (_Materials != null) _Materials.ForEach(new Action<BomMaterial>(x => x.PartID = string.Empty));
                    if (_Processes != null) _Processes.ForEach(new Action<BomProcess>(x => x.PartID = string.Empty));
                }
            }

            #endregion

            #region 表属性
            private List<BomMaterial> _Materials = new List<BomMaterial>();
            public List<BomMaterial> Materials
            {
                get
                {
                    //if (_Materials == null || _Materials.Count <= 0)
                    //{
                    //    LoadBom(this.ID);
                    //}
                    return _Materials;
                }
                set
                { _Materials = value; }
            }
            private List<BomProcess> _Processes = new List<BomProcess>();
            public List<BomProcess> Processes
            {
                get
                {
                    //if (_Processes == null || _Processes.Count <= 0)
                    //{
                    //    LoadBom(this.ID);
                    //}
                    return _Processes;
                }
                set
                { _Processes = value; }
            }
            private List<PartItem> _Parts = null;
            public List<PartItem> Parts
            {
                get
                {
                    //if (_Parts == null || _Parts.Count <= 0)
                    //{
                    //    LoadBom(this.ID);
                    //}
                    return _Parts;
                }
                set
                { _Parts = value; }
            }

            #endregion

            #region 保存


            public bool SaveEstimatePrice()
            {
                try
                {
                    EstimateProductItem bomEI = new EstimateProductItem(Code);
                    Price2 = bomEI.GotPrice();

                    MyData.MyDataParameter[] mps = new MyData.MyDataParameter[]
                    {
                        new MyData.MyDataParameter("@Code",Code),
                        new MyData.MyDataParameter("@Price",Price2, MyData.MyDataParameter.MyDataType.Numeric)
                    };
                    string SQL = @"Update [pbProduct] Set Price2 = @Price Where Code = @Code";
                    MyData.MyCommand mcd = new MyData.MyCommand();
                    return mcd.Execute(SQL, mps);
                }
                catch (Exception ex)
                {
                    return false;
                }
            }


            #endregion

            /// <summary>
            /// 深拷贝
            /// </summary>
            /// <returns></returns>
            public Bom DeepClone()
            {
                return this.Clone<Bom>();
                //using (Stream objectStream = new MemoryStream())
                //{
                //    IFormatter formatter = new BinaryFormatter();
                //    formatter.Serialize(objectStream, this);
                //    objectStream.Seek(0, SeekOrigin.Begin);
                //    return formatter.Deserialize(objectStream) as Bom;
                //}
            }

            public void SetID(int SetID, string SetCode)
            {
                ID = SetID;
                Code = SetCode;
                isNewBom = (SetID == 0);
            }

        }
        #endregion

        #region BOM存储器

        /// <summary>
        /// 重新加载BOM资料
        /// </summary>
        public static void reLoadBomData()
        {
            if (Bom.BomMaterialListCount > 0) Bom.BomMaterialList = null;
            if (Bom.BomProcessListCount > 0) Bom.BomProcessList = null;
            int i = Bom.BomMaterialList.Count();
            i = Bom.BomProcessList.Count();
        }

        #endregion

        #endregion BOM类

        #region 独立工单类


        #endregion

        #endregion 工单/BOM类

        public static class ExportExcel
        {
            public static string Export(MyBase.SendMail sm, IOrderedEnumerable<MyData.MyDataRow> vListGridRow, string[] ColumnFieldNameArray, string[] ColumnCaptionArray, string TableCaption)
            {
                DateTime NowTime = DateTime.Now;
                if (vListGridRow.IsNotEmptySet())
                {
                    string tmpFileName = Path.GetTempFileName();
                    MyRecord.Say(string.Format("tmpFileName = {0}", tmpFileName));
                    Regex rgx = new System.Text.RegularExpressions.Regex(@"(?<=tmp)(.+)(?=\.tmp)");
                    string tmpFileNameLast = rgx.Match(tmpFileName).Value;
                    MyRecord.Say(string.Format("tmpFileNameLast = {0}", tmpFileNameLast));
                    string dName = string.Format("{0}\\TMPExcel", Application.StartupPath);
                    if (!Directory.Exists(dName)) Directory.CreateDirectory(dName);
                    string AttcahFileName = string.Format("{0}\\{1}", dName, string.Format("NOTICE_{0:yyyyMMdd}_tmp{1}.xls", NowTime.Date, tmpFileNameLast));
                    MyRecord.Say(string.Format("AttcahFileName = {0}", AttcahFileName));

                    HSSFWorkbook wb = new HSSFWorkbook();
                    MyRecord.Say(string.Format("输出表格{0}", TableCaption));
                    HSSFSheet st = (HSSFSheet)wb.CreateSheet(TableCaption);
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
                    xCellCaption1.SetCellValue(LCStr(TableCaption));
                    xCellCaption1.CellStyle = cds3;
                    st.AddMergedRegion(new CellRangeAddress(1, 1, 0, ColumnFieldNameArray.Length - 1));

                    //MyRecord.Say(string.Format("输出表格{0}，表格已经创建，建立表头。", caption));

                    HSSFRow rNoteCaption2 = (HSSFRow)st.CreateRow(2);
                    for (int i = 0; i < ColumnFieldNameArray.Length; i++)
                    {
                        ICell xCellCaption = rNoteCaption2.CreateCell(i);
                        xCellCaption.SetCellValue(LCStr(ColumnCaptionArray[i]));
                        xCellCaption.CellStyle = cellNomalStyle;
                        //MyRecord.Say(string.Format("输出表格{0}，标题行，{1}列", caption, LCStr(titles[i])));
                    }

                    MyRecord.Say(string.Format("输出表格{0}，表格已经创建，建立表头完成。", TableCaption));

                    int uu = 3;

                    foreach (var item in vListGridRow)
                    {
                        if (item.IsNotNull())
                        {
                            HSSFRow rNoteCaption3 = (HSSFRow)st.CreateRow(uu);
                            for (int i = 0; i < ColumnFieldNameArray.Length; i++)
                            {
                                string fieldname = ColumnFieldNameArray[i];
                                //MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列", caption, uu, fieldname));
                                ICell xCellCaption = rNoteCaption3.CreateCell(i);
                                if (item.DataRow.Table.Columns.Contains(fieldname))
                                {
                                    //MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列，列存在。", caption, uu, fieldname));
                                    DataColumn dc = item.DataRow.Table.Columns[fieldname];
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
                                    MyRecord.Say(string.Format("输出表格{0}，内容：{1}行，{2}列，没有这个字段。", TableCaption, uu, fieldname));
                                }
                            }
                            uu++;
                        }
                    }

                    MyRecord.Say(string.Format("输出表格{0}完成。", TableCaption));


                    if (File.Exists(AttcahFileName)) File.Delete(AttcahFileName);
                    using (FileStream fs = new FileStream(AttcahFileName, FileMode.Create, FileAccess.Write))
                    {
                        wb.Write(fs);
                        MyRecord.Say("已经保存了。");
                        wb.Close();
                        fs.Close();
                    }

                    if (File.Exists(AttcahFileName))
                    {
                        sm.Attachments.Add(new System.Net.Mail.Attachment(AttcahFileName));
                        MyRecord.Say("加载到附件");
                    }
                    else
                    {
                        MyRecord.Say("没找到附件");
                    }

                    return AttcahFileName;
                }
                return string.Empty ;
            }
        }
    }
}
