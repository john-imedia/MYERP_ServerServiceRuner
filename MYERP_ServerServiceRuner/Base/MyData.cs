using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using MYERP_ServerServiceRuner.Base;

namespace MYERP_ServerServiceRuner
{
    [Serializable]
    public class MyData
    {
        /// <summary>
        /// 运行有返回值的SQL。
        /// </summary>
        [Serializable]
        public class MyDataTable : DataTable
        {
            #region 构造函数，直接会运行执行SQL读取。

            private void ExecuteSQL(string SQL, MyDataParameter[] args, int Interval)
            {
                Type LogT = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType;
                try
                {
                    using (SqlConnection cnn = new SqlConnection(MyBase.ConnectionString))
                    {
                        cnn.Open();
                        try
                        {
                            using (SqlCommand SCMD = new SqlCommand(Convert.ToString(MyConvert.LCDB(SQL)), cnn) { CommandTimeout = 60000 })
                            {
                                try
                                {
                                    SCMD.Parameters.Clear();
                                    if (args != null) SCMD.Parameters.AddRange(args.Select(a => a.sqlparam).ToArray());
                                    SqlDataReader sdr = SCMD.ExecuteReader(CommandBehavior.SingleResult | CommandBehavior.CloseConnection);
                                    Load(sdr);
                                    //st.Commit();
                                    _MyRows = new MyDataRowCollection(Rows);
                                    sdr.Close();
                                }
                                catch (Exception eex)
                                {
                                    //st.Rollback();
                                    throw eex;
                                }
                                finally
                                {
                                    SCMD.Parameters.Clear();
                                }
                            }
                        }
                        finally
                        {
                            cnn.Close();
                        }
                    }
                }
                catch (Exception e)
                {
                    MyRecord.Say(e, SQL, LogT);
                }
            }

            /// <summary>
            /// 从一个DataTable创建
            /// </summary>
            /// <param Name="sourcedatatable"></param>
            public MyDataTable(DataTable sourcedatatable)
            {
                foreach (DataColumn sdi in sourcedatatable.Columns)
                {
                    DataColumn ndi = new DataColumn(sdi.ColumnName, sdi.DataType, sdi.Expression, sdi.ColumnMapping);
                    Columns.Add(ndi);
                }
                Rows.Clear();
                foreach (DataRow sri in sourcedatatable.Rows)
                {
                    ImportRow(sri);
                }
            }

            /// <summary>
            /// 创建后返回SQL语句结果，传入参数。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="SPI">各个参数</param>
            public MyDataTable(string SQL, params MyDataParameter[] SPI)
            {
                ExecuteSQL(SQL, SPI, 3);
            }

            /// <summary>
            /// 创建后返回SQL语句结果，传入参数是语句中{0}的参数。
            /// 会自动调用string.Format来处理SQL语句。注意数据类型。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="strArgs">字符串参数序列</param>
            public MyDataTable(string SQL, params string[] strArgs)
            {
                ExecuteSQL(string.Format(SQL, strArgs), null, 3);
            }

            /// <summary>
            /// 创建的时候，就按照返回SQL语句为准
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            public MyDataTable(string SQL)
            {
                #region 详细的运行方式，测试用

                //SimplyBasic.MyRecorder.Say("准备读取数据库");
                //string ConnectionWord = SimplyBasic.globalVariable.ConnectConfigSet.ConnectionString;
                //SimplyBasic.MyRecorder.Say(string.Format("已经取得连接字符串：{0}\n正在准备连接。",ConnectionWord));
                //SqlConnection SN=new SqlConnection(ConnectionWord);
                //SimplyBasic.MyRecorder.Say(string.Format("已经连接到数据库，下面准备运行SQL语句，SQL语句是：\n{0}",SQL));
                //SqlDataAdapter SQA=new SqlDataAdapter(SQL,SN);
                //SimplyBasic.MyRecorder.Say("SQL语句运行完成，正准备取回数据。");
                //SQA.Fill(this);
                //SimplyBasic.MyRecorder.Say(string.Format("结果集已经取回，取回了：{0}", this.Rows.Count));
                #endregion 详细的运行方式，测试用

                ExecuteSQL(SQL, null, 3);
            }

            /// <summary>
            /// 创建后返回SQL语句结果，传入参数。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="SPI">各个参数</param>
            public MyDataTable(string SQL, int Interval, params MyDataParameter[] SPI)
            {
                ExecuteSQL(SQL, SPI, Interval);
            }

            /// <summary>
            /// 创建后返回SQL语句结果，传入参数是语句中{0}的参数。
            /// 会自动调用string.Format来处理SQL语句。注意数据类型。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="strArgs">字符串参数序列</param>
            public MyDataTable(string SQL, int Interval, params string[] strArgs)
            {
                ExecuteSQL(string.Format(SQL, strArgs), null, Interval);
            }

            /// <summary>
            /// 创建的时候，就按照返回SQL语句为准
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            public MyDataTable(string SQL, int Interval)
            {
                #region 详细的运行方式，测试用

                //SimplyBasic.MyRecorder.Say("准备读取数据库");
                //string ConnectionWord = SimplyBasic.globalVariable.ConnectConfigSet.ConnectionString;
                //SimplyBasic.MyRecorder.Say(string.Format("已经取得连接字符串：{0}\n正在准备连接。",ConnectionWord));
                //SqlConnection SN=new SqlConnection(ConnectionWord);
                //SimplyBasic.MyRecorder.Say(string.Format("已经连接到数据库，下面准备运行SQL语句，SQL语句是：\n{0}",SQL));
                //SqlDataAdapter SQA=new SqlDataAdapter(SQL,SN);
                //SimplyBasic.MyRecorder.Say("SQL语句运行完成，正准备取回数据。");
                //SQA.Fill(this);
                //SimplyBasic.MyRecorder.Say(string.Format("结果集已经取回，取回了：{0}", this.Rows.Count));
                #endregion 详细的运行方式，测试用

                ExecuteSQL(SQL, null, Interval);
            }

            #region 隐藏不可以被引用的构造函数

            private MyDataTable()
            {
            }

            private MyDataTable(string tableName, string tableNamespace)
                : base(tableName, tableNamespace)
            {
            }

            private MyDataTable(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context)
                : base(info, context)
            {
            }

            #endregion 隐藏不可以被引用的构造函数

            #endregion 构造函数，直接会运行执行SQL读取。

            #region 属性
            private MyDataRowCollection _MyRows;

            public MyDataRowCollection MyRows
            {
                get
                {
                    //这样就可以把DateTable直接赋值过来了。
                    if (_MyRows == null)
                    {
                        _MyRows = new MyDataRowCollection(Rows);
                    }

                    return _MyRows;
                }
            }



            #endregion 属性
        }

        /// <summary>
        /// 处理数据行，转换简繁体
        /// </summary>
        [Serializable]
        public class MyDataRow : IEnumerable, IEnumerator
        {
            #region 存储器
            private readonly DataRow _RowData;
            private int CurID = -1;
            #endregion 存储器
            /// <summary>
            /// 直接的后台数据
            /// </summary>
            public DataRow DataRow
            {
                get
                {
                    return _RowData;
                }
            }

            public bool Contains(string ColumnName)
            {
                return _RowData.Table.Columns.Contains(ColumnName);
            }

            #region 构造器

            public MyDataRow(DataRow DataRowForConvert)
            {
                _RowData = DataRowForConvert;
            }

            #endregion 构造器

            #region 索引器

            public object this[string columnName]
            {
                get
                {
                    if (_RowData.Table.Columns.Contains(columnName))
                    {
                        return this[_RowData.Table.Columns[columnName]];
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            public object this[DataColumn column]
            {
                get
                {
                    string name = column.ColumnName;
                    Type colType = column.DataType;
                    object xValue = _RowData[column];
                    if (xValue != DBNull.Value)
                    {
                        if (colType == typeof(string))
                            return StringValue(name);
                        else if (colType == typeof(bool))
                            return BooleanValue(name);
                        else if (colType == typeof(Guid))
                            return Value<Guid>(name);
                        else if (colType == typeof(DateTime))
                            return Value<DateTime>(name);
                        else
                            return xValue;
                    }
                    else
                    {
                        if (colType == typeof(bool))
                            return false;
                        else
                            return null;
                    }
                }
            }

            public object this[int columnIndex]
            {
                get
                {
                    if (columnIndex > 0 && columnIndex < _RowData.Table.Columns.Count)
                    {
                        return this[_RowData.Table.Columns[columnIndex]];
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            public TSource Value<TSource>(string ColumnName)
            {
                try
                {
                    if (_RowData.Table.Columns.Contains(ColumnName))
                    {
                        DataColumn colItem = _RowData.Table.Columns[ColumnName];
                        object xValue = _RowData[colItem];
                        Type SourceType = typeof(TSource);
                        Type ColumnType = colItem.DataType;
                        if (xValue != DBNull.Value)
                        {
                            if (SourceType == typeof(Guid))
                            {
                                string x = xValue.ConvertTo<string>(string.Empty, true);
                                Guid r;
                                if (Guid.TryParse(x, out r))
                                {
                                    return (TSource)(object)r;
                                }
                                else
                                {
                                    return (TSource)(object)Guid.Empty;
                                }
                            }
                            else if (SourceType == typeof(string))
                            {
                                string x = xValue.ConvertTo<string>(string.Empty, true) ?? string.Empty;
                                return (TSource)(object)MyConvert.DBLC(x);
                            }
                            else if (SourceType == typeof(bool))
                            {
                                if (colItem.DataType == typeof(bool))
                                {
                                    return (TSource)(object)xValue.ConvertTo<bool>(false, true);
                                }
                                else if (xValue.ToString().IsNumeric())
                                {
                                    return (TSource)(object)(xValue.ConvertTo<double>(0, true) != 0);
                                }
                            }
                            else if (SourceType == typeof(int))
                            {
                                double dValue = 0;
                                string dValueStr = xValue.ConvertTo<string>("0", true);
                                if (!double.TryParse(dValueStr, out dValue)) dValue = 0;
                                return (TSource)(object)Convert.ToInt32(Math.Round(dValue));
                            }
                            else
                            {
                                return xValue.ConvertTo<TSource>(default(TSource), true);
                            }
                        }
                        else if (ColumnType == typeof(string))
                        {
                            if (SourceType == typeof(string)) return (TSource)(object)string.Empty;
                        }
                    }
                }
                catch (Exception e)
                {
                    MyRecord.Say(e);
                }

                return default(TSource);
            }

            public string Value(string ColumnName)
            {
                return Value<string>(ColumnName);
            }

            public string StringValue(string ColumnName)
            {
                return Value<string>(ColumnName);
            }

            public int IntValue(string ColumnName)
            {
                return Value<int>(ColumnName);
            }

            public bool BooleanValue(string ColumnName)
            {
                return Value<bool>(ColumnName);
            }

            public Guid GuidValue(string ColumnName)
            {
                return Value<Guid>(ColumnName);
            }

            public DateTime DateTimeValue(string ColumnName)
            {
                return Value<DateTime>(ColumnName);
            }

            #endregion 索引器

            #region 实现ENUM接口

            public IEnumerator GetEnumerator()
            {
                return (IEnumerator)this;
            }

            public void Reset()
            {
                CurID = -1;
            }

            public bool MoveNext()
            {
                CurID++;
                return CurID >= 0 && CurID < _RowData.ItemArray.Length;
            }

            public object Current
            {
                get
                {
                    if (CurID >= 0 && CurID < _RowData.ItemArray.Length)
                    {
                        return MyConvert.DBLC(_RowData[CurID].ToString());
                    }
                    else
                    {
                        return null;
                    }
                }
                set
                {
                    if (CurID >= 0 && CurID < _RowData.ItemArray.Length)
                    {
                        _RowData[CurID] = value;
                    }
                }
            }

            public int Length
            {
                get
                {
                    return _RowData.ItemArray.Length;
                }
            }

            #endregion 实现ENUM接口
        }

        /// <summary>
        /// 很多自定义的数据行，可以直接实现带着简繁体和去掉DBNull后的行。
        /// </summary>
        [Serializable]
        public class MyDataRowCollection : MyBase.MyEnumerable<MyDataRow>, IDisposable
        {
            #region 存储器
            private DataRowCollection _DataRows;
            #endregion 存储器

            public DataRowCollection DataRows
            {
                get
                {
                    return _DataRows;
                }
            }

            /// <summary>
            /// 第一行或者唯一的一行。
            /// </summary>
            public MyDataRow FirstOrOnlyRow
            {
                get
                {
                    if (_data != null && _data.Count() > 0)
                        return _data.FirstOrDefault();
                    else
                        return null;
                }
            }

            /// <summary>
            /// 是否有数据行
            /// </summary>
            public bool HaveRow
            {
                get
                {
                    if (_DataRows != null)
                        return (_DataRows.Count > 0);
                    else
                        return false;
                }
            }


            #region 构造器

            public MyDataRowCollection(DataRowCollection xRows)
            {
                _DataRows = xRows;
                LoadDataSource();
            }

            public MyDataRowCollection(DataRow[] ArrayDataRow)
            {
                if (ArrayDataRow.Length > 0)
                {
                    using (DataTable dt = new DataTable())
                    {
                        DataColumnCollection dc = ArrayDataRow[0].Table.Columns;
                        DataColumn[] xar = new DataColumn[dc.Count];
                        int i = -1;
                        foreach (DataColumn dCol in dc)
                        {
                            i++;
                            xar[i] = new DataColumn(dCol.ColumnName, dCol.DataType, dCol.Expression) { AllowDBNull = true };
                        }
                        dt.Columns.AddRange(xar);
                        foreach (DataRow xaw in ArrayDataRow)
                        {
                            DataRow xxa = dt.NewRow();
                            foreach (DataColumn dCol in dc)
                            {
                                xxa[dCol.ColumnName] = xaw[dCol];
                            }
                            dt.Rows.Add(xxa);
                        }
                        _DataRows = dt.Rows;
                        LoadDataSource();
                    }
                }
            }

            public MyDataRowCollection(DataTable xTable)
            {
                if (xTable != null)
                {
                    _DataRows = xTable.Rows;
                    LoadDataSource();
                }
            }

            private void LoadDataSource()
            {
                var v = from DataRow a in _DataRows
                        select new MyDataRow(a);
                _data = v.ToList();
            }


            #endregion 构造器


            #region 实现接口
            public new int Count
            {
                get
                {
                    return Count();
                }
            }

            void IDisposable.Dispose()
            {
            }

            #endregion 实现接口
        }

        [Serializable]
        /// <summary>
        /// 运行SQL语句命令。
        /// </summary>
        public class MyCommand
        {
            #region 存储器
            private MyDataCmdCollection _myCmdColl;
            #endregion 存储器

            #region 运行无有返回值的SQL语句

            #region 各种重载，方便使用

            /// <summary>
            /// 运行一条SQL语句
            /// 注意：这将覆盖SQLCmdColl属性，将会使SQLCmdColl属性变成新的只有一条SQL的序列。
            /// 出错会回滚操作。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <returns>会返回是否出错。</returns>
            public bool Execute(string SQL)
            {
                SQLCmdColl = new MyDataCmdCollection();
                SQLCmdColl.Add(SQL);
                return Execute();
            }

            /// <summary>
            /// 运行一条SQL语句
            /// 注意：这将覆盖SQLCmdColl属性，将会使SQLCmdColl属性变成新的只有一条SQL的序列。
            /// 出错会回滚操作。可以很多参数
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="Parameters">很多参数，或者一个数组</param>
            /// <returns></returns>
            public bool Execute(string SQL, params MyDataParameter[] Parameters)
            {
                SQLCmdColl = new MyDataCmdCollection();
                SQLCmdColl.Add(SQL, "Main", Parameters);
                return Execute();
            }

            /// <summary>
            /// 运行一条SQL语句
            /// 注意：这将覆盖SQLCmdColl属性，将会使SQLCmdColl属性变成新的只有一条SQL的序列。
            /// 出错会回滚操作。固定一个参数。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="param0">固定一个参数</param>
            /// <returns></returns>
            public bool Execute(string SQL, SqlParameter param0)
            {
                if (param0 != null)
                {
                    return Execute(SQL, param0);
                }
                else
                {
                    return Execute(SQL);
                }
            }

            /// <summary>
            /// 运行一条SQL语句
            /// 注意：这将覆盖SQLCmdColl属性，将会使SQLCmdColl属性变成新的只有一条SQL的序列。
            /// 出错会回滚操作。固定两个参数。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="param0">第一个参数</param>
            /// <param Name="param1">第二个参数</param>
            /// <returns></returns>
            public bool Execute(string SQL, MyDataParameter param0, MyDataParameter param1)
            {
                MyDataParameter[] ps = new MyDataParameter[2] { param0, param1 };
                return Execute(SQL, ps);
            }

            /// <summary>
            /// 运行一条SQL语句
            /// 注意：这将覆盖SQLCmdColl属性，将会使SQLCmdColl属性变成新的只有一条SQL的序列。
            /// 出错会回滚操作。固定三个参数
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="param0">第一个参数</param>
            /// <param Name="param1">第二个参数</param>
            /// <param Name="param2">第三个参数</param>
            /// <returns></returns>
            public bool Execute(string SQL, MyDataParameter param0, MyDataParameter param1, MyDataParameter param2)
            {
                MyDataParameter[] ps = new MyDataParameter[3] { param0, param1, param2 };
                return Execute(SQL, ps);
            }

            /// <summary>
            /// 运行SQLCmdColl属性中的SQL语句。
            /// 这些SQL语句是不返回值的。
            /// SQLList会覆盖SQLCmdColl属性。
            /// 运行语句是有事务的，出错会回滚操作。
            /// </summary>
            /// <param Name="SQLList">设置要运行的语句序列，会覆盖SQLCmdColl属性</param>
            /// <returns>运行是否正确，不正确将RollBack</returns>
            public bool Execute(MyDataCmdCollection SQLList)
            {
                SQLCmdColl = SQLList;
                return Execute();
            }

            #endregion 各种重载，方便使用

            /// <summary>
            /// 运行Add进入或者，SQLCmdColl属性中的SQL语句。
            /// 这些语句都是不返回值的！
            /// 运行有事务，出错会回滚。
            /// </summary>
            /// <returns>出错会回滚</returns>
            public bool Execute()
            {
                bool ExecOK = false;
                Type LogT = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType;
                try
                {
                    using (SqlConnection sn = new SqlConnection(MyBase.ConnectionString))
                    {
                        sn.Open();
                        SqlTransaction st = null;
                        if (SQLCmdColl.Count > 1) st = sn.BeginTransaction();
                        using (SqlCommand sc = new SqlCommand() { CommandType = CommandType.Text, Connection = sn, CommandTimeout = 60000 })
                        {
                            try
                            {
                                if (st != null) sc.Transaction = st;
                                SQLCmdColl.Reset();
                                int i = 0, l = SQLCmdColl.Count;
                                foreach (MyDataCmdItem mc in SQLCmdColl)
                                {
                                    try
                                    {
                                        if (mc.SQL.Trim().Length <= 0)
                                        {
                                            throw new Exception(string.Format("没有给SQL设置语句！,第{0}句。Key={1}", mc.index, mc.key));
                                        }
                                        i++;
                                        sc.CommandText = MyConvert.LCDB(mc.SQL).ToString(); //设置SQL语句，要进行简繁体转换；
                                        sc.Parameters.Clear();
                                        if (mc.SQLParamCollection != null)
                                        {
                                            sc.Parameters.Clear();
                                            sc.Parameters.AddRange(mc.SQLParamCollection);
                                        }
                                        mc.EffectRows = sc.ExecuteNonQuery(); //运行SQL语句。
                                        mc.ExecuteOK = true;
                                    }
                                    catch (Exception e)
                                    {
                                        MyRecord.Say(e, string.Format("SQLKey={0}", mc.key), LogT);
                                        mc.ExecuteOK = false;
                                        mc.ExecuteExp = e;
                                        throw e;
                                    }
                                }
                                Thread.Sleep(100);
                                if (st != null) st.Commit();
                                ExecOK = true;
                            }
                            catch (Exception e)
                            {
                                if (st != null) st.Rollback();
                                throw e;
                            }
                            finally
                            {
                                sc.Parameters.Clear();
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    MyRecord.Say(e, "运行批量SQL出错。", LogT);
                    ExecOK = false;
                }
                return ExecOK;
            }

            #endregion 运行无有返回值的SQL语句

            #region 多条SQL语句
            /// <summary>
            /// SQL语句序列
            /// </summary>
            public MyDataCmdCollection SQLCmdColl
            {
                get
                {
                    if (_myCmdColl == null)
                    {
                        _myCmdColl = new MyDataCmdCollection();
                    }
                    return _myCmdColl;
                }
                set
                {
                    _myCmdColl = value;
                }
            }

            #region 加载SQL语句

            /// <summary>
            /// 加一个SQL语句
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="Key">KEY</param>
            /// <param Name="args">参数序列</param>
            public void Add(string SQL, string Key, params MyDataParameter[] args)
            {
                SQLCmdColl.Add(SQL, Key, args);
            }

            /// <summary>
            /// 加一个SQL语句
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="Key">KEY</param>
            public void Add(string SQL, string Key)
            {
                SQLCmdColl.Add(SQL, Key);
            }

            /// <summary>
            /// 加入一个SQL语句，没有KEY，出错时将不知道哪里错误。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            public void Add(string SQL)
            {
                SQLCmdColl.Add(SQL);
            }

            /// <summary>
            /// 加入一个SQL语句对象
            /// </summary>
            /// <param Name="SQLItem">SQL语句对象</param>
            public void Add(MyDataCmdItem SQLItem)
            {
                SQLCmdColl.Add(SQLItem);
            }

            #endregion 加载SQL语句

            #endregion 多条SQL语句

            #region 多条SQL语句存储类
            [Serializable]
            /// <summary>
            /// 要运行的一条SQL命令。
            /// </summary>
            public class MyDataCmdItem
            {
                #region 存储器

                public string SQL { get; set; }

                public System.Exception ExecuteExp { get; set; }

                public bool ExecuteOK { get; set; }

                public string key { get; set; }

                public int index { get; set; }

                public int EffectRows { get; set; }

                public SqlParameter[] SQLParamCollection { get; set; }

                public MyDataParameter[] MyDataParamCollection { get; set; }

                #endregion 存储器

                #region 构造器

                public MyDataCmdItem()
                {
                    index = -1;
                }

                public MyDataCmdItem(string sql)
                {
                    SQL = sql;
                    key = string.Empty;
                    index = -1;
                }

                public MyDataCmdItem(string sql, string Key)
                {
                    SQL = sql;
                    key = Key;
                    index = -1;
                }

                public MyDataCmdItem(string sql, string Key, params MyDataParameter[] args)
                {
                    SQL = sql;
                    key = Key;
                    index = -1;
                    MyDataParamCollection = args;
                    //Param = new SqlParameter[args.Length];
                    //for (int i = 0; i < args.Length; i++)
                    //{
                    //    Param[i] = args[i].sqlparam;
                    //}
                    if (args.IsNotNull() && args.Length > 0)
                    {
                        var v = from a in args
                                select a.sqlparam;
                        SQLParamCollection = v.ToArray();
                    }
                }

                public MyDataCmdItem(string sql, params MyDataParameter[] args)
                {
                    SQL = sql;
                    index = -1;
                    MyDataParamCollection = args;
                    if (args.IsNotNull() && args.Length > 0)
                    {
                        var v = from a in args
                                select a.sqlparam;
                        SQLParamCollection = v.ToArray();
                    }
                }

                #endregion 构造器
            }


            [Serializable]
            /// <summary>
            /// 保持要运行的SQL集合。
            /// </summary>
            public class MyDataCmdCollection : MyBase.MyEnumerable<MyDataCmdItem>
            {
                #region 构造器

                public MyDataCmdCollection()
                {
                }

                #endregion 构造器

                #region 加减元素

                public MyDataCmdItem Add(string SQL, string Key, params MyDataParameter[] args)
                {
                    MyDataCmdItem cItem = new MyDataCmdItem(SQL, Key, args);
                    Add(cItem);
                    return cItem;
                }

                public MyDataCmdItem Add(string SQL, string Key)
                {
                    MyDataCmdItem cItem = new MyDataCmdItem(SQL, Key);
                    Add(cItem);
                    return cItem;
                }

                public MyDataCmdItem Add(string SQL)
                {
                    MyDataCmdItem cItem = new MyDataCmdItem(SQL);
                    Add(cItem);
                    return cItem;
                }

                public MyDataCmdItem Add(string SQL, params MyDataParameter[] args)
                {
                    MyDataCmdItem cItem = new MyDataCmdItem(SQL, args);
                    Add(cItem);
                    return cItem;
                }

                public void Remove(string KEY)
                {
                    try
                    {
                        var v = from a in _data
                                where a.key == KEY
                                select a;
                        if (v.IsNotNull() && v.Count() > 0)
                        {
                            RemoveItem(v.FirstOrDefault());
                        }
                    }
                    catch (Exception ex)
                    {
                        MyRecord.Say(ex);
                    }
                }
                #endregion 加减元素

                public new int Count
                {
                    get
                    {
                        return Count();
                    }
                }
            }

            #endregion 多条SQL语句存储类

            #region 运行一个返回一个值的一条SQL语句

            /// <summary>
            /// 从数据库返回一个值
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="paramArgs">参数序列</param>
            /// <returns>一个值，注意要（类型）objec强制类型转换</returns>
            public static object ExecuteScalar(string SQL, params MyData.MyDataParameter[] paramArgs)
            {
                Type LogT = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType;
                try
                {
                    SqlConnection cnn = new SqlConnection(MyBase.ConnectionString);
                    cnn.Open();
                    SQL = MyConvert.LCDB(SQL).ToString(); //进行繁简转换。
                    SqlTransaction st = cnn.BeginTransaction();
                    SqlCommand SCMD = new SqlCommand(SQL, cnn, st);
                    SqlParameter[] sqlpa = null;
                    if (paramArgs != null)
                    {
                        sqlpa = new SqlParameter[paramArgs.Length];
                        for (int i = 0; i < paramArgs.Length; i++)
                        {
                            sqlpa[i] = paramArgs[i].sqlparam;
                        }
                        SCMD.Parameters.Clear();
                        SCMD.Parameters.AddRange(sqlpa);
                    }
                    SCMD.CommandText = SQL;
                    SCMD.CommandType = CommandType.Text;
                    try
                    {
                        object _R = SCMD.ExecuteScalar();
                        st.Commit();
                        return _R;
                    }
                    catch (Exception e)
                    {
                        st.Rollback();
                        MyRecord.Say(e, SQL, LogT);
                        return null;
                    }
                    finally
                    {
                        SCMD.Parameters.Clear();
                        st.Dispose();
                        SCMD.Dispose();
                        cnn.Close();
                        cnn.Dispose();
                    }
                }
                catch (Exception e)
                {
                    MyRecord.Say(e, SQL, LogT);
                    return null;
                }
            }

            /// <summary>
            /// 从数据库返回一个值
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <returns>一个值，注意要（类型）objec强制类型转换</returns>
            public static object ExecuteScalar(string SQL)
            {
                return ExecuteScalar(SQL, (MyData.MyDataParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值，支持{0}参数，内部用string.Format自动格式化。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="strArgs">字符串参数序列</param>
            /// <returns>一个值，注意要（类型）objec强制类型转换</returns>
            public static object ExecuteScalar(string SQL, params string[] strArgs)
            {
                return ExecuteScalar(string.Format(SQL, strArgs), (MyData.MyDataParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="strArg0">第一个参数</param>
            /// <returns>一个值，注意要（类型）objec强制类型转换</returns>
            public static object ExecuteScalar(string SQL, string strArg0)
            {
                return ExecuteScalar(string.Format(SQL, strArg0), (MyData.MyDataParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="strArg0">第一个参数</param>
            /// <param Name="strArg1">第二个参数</param>
            /// <returns></returns>
            public static object ExecuteScalar(string SQL, string strArg0, string strArg1)
            {
                return ExecuteScalar(string.Format(SQL, strArg0, strArg1), (MyData.MyDataParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="strArg0">第一个参数</param>
            /// <param Name="strArg1">第二个参数</param>
            /// <param Name="strArg2">第三个参数</param>
            /// <returns></returns>
            public static object ExecuteScalar(string SQL, string strArg0, string strArg1, string strArg2)
            {
                return ExecuteScalar(string.Format(SQL, strArg0, strArg1, strArg2), (MyData.MyDataParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值,用SQL参数的。
            /// </summary>
            /// <param Name="SQL">SQL语句</param>
            /// <param Name="paramArg0">第一个参数</param>
            /// <returns></returns>
            public static object ExecuteScalar(string SQL, MyData.MyDataParameter paramArg0)
            {
                MyData.MyDataParameter[] mp = new MyDataParameter[1];
                mp[0] = paramArg0;
                return ExecuteScalar(SQL, mp);
            }

            #endregion 运行一个返回一个值的一条SQL语句
        }

        /// <summary>
        /// SQL语句参数类。
        /// </summary>
        [Serializable]
        public class MyDataParameter
        {
            [Serializable]
            public enum MyDataType : uint
            {
                Boolean,
                DateTime,
                Int,
                Numeric,
                String,
                UniqueIdentifier
            }

            private object _Value = null;
            private string _Name = string.Empty;
            private MyDataType _MyDbType = MyDataType.String;
            public SqlParameter sqlparam;

            public MyDataParameter()
            {
                sqlparam = new SqlParameter();
                MyDbType = MyDataType.String;
                Value = null;
                Name = "default";
            }

            /// <summary>
            /// 默认Type
            /// 是String；
            /// </summary>
            /// <param Name="Name"></param>
            /// <param Name="value"></param>
            public MyDataParameter(string name, object value)
            {
                sqlparam = new SqlParameter();
                MyDbType = MyDataType.String;
                Value = value;
                Name = name;

                //sqlparam = new SqlParameter(_Name, _Value);
            }

            public MyDataParameter(string name, object value, MyDataType mtype)
            {
                sqlparam = new SqlParameter();
                MyDbType = mtype;
                Value = value;
                Name = name;

                //sqlparam = new SqlParameter(_Name, _Value) { SqlDbType = getSqlDbType(mtype) };
            }

            public string Name
            {
                get
                {
                    return _Name;
                }
                set
                {
                    _Name = value;
                    sqlparam.ParameterName = value;
                }
            }

            public object Value
            {
                get
                {
                    return MyConvert.DBLC(_Value);
                }
                set
                {
                    if (Convert.IsDBNull(value))
                    {
                        _Value = DBNull.Value;
                    }
                    else
                    {
                        if (_MyDbType == MyDataType.Numeric || _MyDbType == MyDataType.Int)
                        {
                            if (value == null)
                            {
                                _Value = 0;
                            }
                            else
                            {
                                if (value.ToString() == string.Empty)
                                {
                                    _Value = 0;
                                }
                                else
                                {
                                    if (value.GetType() == typeof(string))
                                    {
                                        if (Convert.ToString(value).Trim() == string.Empty)
                                        {
                                            _Value = 0;
                                        }
                                        else
                                        {
                                            if (_MyDbType == MyDataType.Int)
                                            {
                                                int v;
                                                if (int.TryParse(Convert.ToString(value), out v))
                                                {
                                                    _Value = v;
                                                }
                                                else
                                                {
                                                    _Value = 0;
                                                }
                                            }
                                            else
                                            {
                                                if (_MyDbType == MyDataType.Numeric)
                                                {
                                                    double v;
                                                    if (double.TryParse(Convert.ToString(value), out v))
                                                    {
                                                        _Value = v;
                                                    }
                                                    else
                                                    {
                                                        _Value = 0;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (_MyDbType == MyDataType.Int)
                                        {
                                            _Value = Convert.ToInt32(value);
                                        }
                                        else
                                        {
                                            if (_MyDbType == MyDataType.Numeric)
                                            {
                                                double tValue = Convert.ToDouble(value);
                                                if ((tValue * 100) - Math.Floor(tValue * 100) == 0)    //判断两位小数
                                                {
                                                    sqlparam.SqlDbType = SqlDbType.Money;
                                                    _Value = Convert.ToDecimal(Math.Round(tValue, 2));
                                                }
                                                else if ((tValue * 1000000) - Math.Floor(tValue * 1000000) == 0)    //判断4位小数
                                                {
                                                    sqlparam.SqlDbType = SqlDbType.Decimal;
                                                    _Value = Convert.ToSingle(Math.Round(tValue, 12));
                                                }
                                                else
                                                {
                                                    _Value = tValue;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else if (_MyDbType == MyDataType.Boolean)
                        {
                            if (value == null)
                            {
                                _Value = false;
                            }
                            else
                            {
                                if (value.GetType() == typeof(bool))
                                {
                                    _Value = (bool)value ? 1 : 0;
                                }
                                else
                                {
                                    if (value.GetType() == typeof(string))
                                    {
                                        _Value = value.ToString().ToLower() == "true";
                                    }
                                    else
                                    {
                                        try
                                        {
                                            _Value = Convert.ToBoolean(value);
                                        }
                                        catch
                                        {
                                            _Value = Convert.ToByte(value) == 1;
                                        }
                                    }
                                }
                            }
                        }
                        else if (_MyDbType == MyDataType.DateTime)
                        {
                            DateTime dv;
                            if (value == null)
                            {
                                _Value = DBNull.Value;
                            }
                            else
                            {
                                if (DateTime.TryParse(Convert.ToString(value), out dv))
                                {
                                    if (dv == DateTime.MinValue || dv == DateTime.MaxValue)
                                    {
                                        _Value = DBNull.Value;
                                    }
                                    else
                                    {
                                        _Value = Convert.ToDateTime(value);
                                    }
                                }
                                else
                                {
                                    _Value = DBNull.Value;
                                }
                            }
                        }
                        else if (_MyDbType == MyDataType.UniqueIdentifier)
                        {
                            if (value == null)
                            {
                                _Value = DBNull.Value;
                            }
                            else
                            {
                                Guid r;
                                if (Guid.TryParse(value.ToString(), out r))
                                    _Value = r;
                                else
                                    _Value = DBNull.Value;
                            }
                        }
                        else
                        {
                            if (value == null)
                            {
                                _Value = DBNull.Value;
                            }
                            else
                            {
                                _Value = MyConvert.LCDB(value);
                            }
                        }
                    }

                    sqlparam.Value = _Value;
                }
            }

            public MyDataType MyDbType
            {
                get
                {
                    return _MyDbType;
                }
                set
                {
                    _MyDbType = value;
                    sqlparam.SqlDbType = getSqlDbType(value);
                    Value = _Value;
                }
            }

            private static SqlDbType getSqlDbType(MyDataType mType)
            {
                switch (mType)
                {
                    case MyDataType.String:
                        return SqlDbType.NVarChar;
                    case MyDataType.Int:
                        return SqlDbType.Int;
                    case MyDataType.Numeric:
                        return SqlDbType.Real;
                    case MyDataType.DateTime:
                        return SqlDbType.DateTime;
                    case MyDataType.Boolean:
                        return SqlDbType.Bit;
                    case MyDataType.UniqueIdentifier:
                        return SqlDbType.UniqueIdentifier;
                    default:
                        return SqlDbType.NVarChar;
                }
            }
        }

        #region 数据返回的固定信息。

        /// <summary>
        /// 返回数据库时间,此时间是从数据库返回的。要查询数据库。
        /// </summary>
        public static DateTime SysDate
        {
            get
            {
                try
                {
                    return (DateTime)MyCommand.ExecuteScalar("Select GetDate()");
                }
                catch
                {
                    return DateTime.Now;
                }
            }
        }

        /// <summary>
        /// 返回计算的当前时间。不会有数据参数
        /// </summary>
        public static DateTime Now
        {
            get
            {
                return SysDate;
            }
        }

        #endregion 数据返回的固定信息。
    }
}