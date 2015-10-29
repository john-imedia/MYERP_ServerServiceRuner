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
    public class MyData
    {
        /// <summary>
        /// 运行有返回值的SQL。
        /// </summary>
        [Serializable]
        public class MyDataTable : DataTable
        {
            #region 构造函数，直接会运行执行SQL读取。

            private void ExecuteSQL(string SQL, MyParameter[] args, int Interval)
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
                                    SqlDataReader sdr = SCMD.ExecuteReader(CommandBehavior.SingleResult);
                                    Load(sdr);
                                    sdr.Close();
                                    _MyRows = new MyDataRowColl(Rows);
                                }
                                catch (Exception eex)
                                {
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
            /// <param name="sourcedatatable"></param>
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
            /// <param name="SQL">SQL语句</param>
            /// <param name="SPI">各个参数</param>
            public MyDataTable(string SQL, params MyParameter[] SPI)
            {
                ExecuteSQL(SQL, SPI, 3);
            }

            /// <summary>
            /// 创建后返回SQL语句结果，传入参数是语句中{0}的参数。
            /// 会自动调用string.Format来处理SQL语句。注意数据类型。
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="strArgs">字符串参数序列</param>
            public MyDataTable(string SQL, params string[] strArgs)
            {
                ExecuteSQL(string.Format(SQL, strArgs), null, 3);
            }

            /// <summary>
            /// 创建的时候，就按照返回SQL语句为准
            /// </summary>
            /// <param name="SQL">SQL语句</param>
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
            /// <param name="SQL">SQL语句</param>
            /// <param name="SPI">各个参数</param>
            public MyDataTable(string SQL, int Interval, params MyParameter[] SPI)
            {
                ExecuteSQL(SQL, SPI, Interval);
            }

            /// <summary>
            /// 创建后返回SQL语句结果，传入参数是语句中{0}的参数。
            /// 会自动调用string.Format来处理SQL语句。注意数据类型。
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="strArgs">字符串参数序列</param>
            public MyDataTable(string SQL, int Interval, params string[] strArgs)
            {
                ExecuteSQL(string.Format(SQL, strArgs), null, Interval);
            }

            /// <summary>
            /// 创建的时候，就按照返回SQL语句为准
            /// </summary>
            /// <param name="SQL">SQL语句</param>
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
            private MyDataRowColl _MyRows;

            public MyDataRowColl MyRows
            {
                get
                {
                    //这样就可以把DateTable直接赋值过来了。
                    if (_MyRows == null)
                    {
                        _MyRows = new MyDataRowColl(Rows);
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
            private readonly DataRow RowData;
            private int CurID = -1;
            #endregion 存储器

            public DataRow DataRow
            {
                get
                {
                    return RowData;
                }
            }

            public bool Contains(string ColumnName)
            {
                return RowData.Table.Columns.Contains(ColumnName);
            }

            #region 构造器

            public MyDataRow(DataRow DataRowForConvert)
            {
                RowData = DataRowForConvert;
            }

            #endregion 构造器

            #region 索引器

            public object this[string columnName]
            {
                get
                {
                    try
                    {
                        if (RowData.Table.Columns.Contains(columnName))
                        {
                            return this[RowData.Table.Columns[columnName]];
                        }
                        else
                        {
                            return null;
                        }
                    }
                    catch (Exception e)
                    {
                        MyRecord.Say(e);
                        return null;
                    }
                }
                set
                {
                    try
                    {
                        if (!RowData.Table.Columns.Contains(columnName))
                        {
                            return;
                        }
                        this[RowData.Table.Columns[columnName]] = value;
                    }
                    catch (Exception e)
                    {
                        MyRecord.Say(e);
                    }
                }
            }

            public object this[DataColumn column]
            {
                get
                {
                    try
                    {
                        if (RowData[column] == DBNull.Value)
                        {
                            if (column.DataType == typeof(bool))
                            {
                                return false;
                            }
                            else
                            {
                                return null;
                            }
                        }
                        else
                        {
                            if (column.DataType == typeof(string))
                            {
                                return MyConvert.DBLC(RowData[column]);
                            }
                            else
                            {
                                try
                                {
                                    if (column.DataType == typeof(byte) || column.DataType == typeof(Int16) || column.DataType == typeof(int))
                                        return Convert.ToInt32(RowData[column]);
                                    else if (column.DataType == typeof(bool))
                                        return Convert.ToBoolean(RowData[column]);
                                    else
                                        return RowData[column];
                                }
                                catch (Exception ex)
                                {
                                    MyRecord.Say(ex);
                                    return RowData[column];
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        MyRecord.Say(e);
                        return null;
                    }
                }
                set
                {
                    try
                    {
                        if (value == null)
                        {
                            RowData[column] = DBNull.Value;
                        }
                        else
                        {
                            if (column.DataType == typeof(string))
                            {
                                RowData[column] = MyConvert.LCDB(value);
                            }
                            else
                            {
                                RowData[column] = value;
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        MyRecord.Say(e);
                    }
                }
            }

            public object this[int columnIndex]
            {
                get
                {
                    try
                    {
                        if (columnIndex > 0 && columnIndex < RowData.Table.Columns.Count)
                        {
                            return this[RowData.Table.Columns[columnIndex]];
                        }
                        else
                        {
                            return null;
                        }
                    }
                    catch (Exception e)
                    {
                        MyRecord.Say(e);
                        return null;
                    }
                }
                set
                {
                    try
                    {
                        if (columnIndex > 0 && columnIndex < RowData.Table.Columns.Count)
                        {
                            this[RowData.Table.Columns[columnIndex]] = value;
                        }
                        else
                        {
                            return;
                        }
                    }
                    catch (Exception e)
                    {
                        MyRecord.Say(e);
                    }
                }
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
                return CurID >= 0 && CurID < RowData.ItemArray.Length;
            }

            public object Current
            {
                get
                {
                    if (CurID >= 0 && CurID < RowData.ItemArray.Length)
                    {
                        return MyConvert.DBLC(RowData[CurID].ToString());
                    }
                    else
                    {
                        return null;
                    }
                }
                set
                {
                    if (CurID >= 0 && CurID < RowData.ItemArray.Length)
                    {
                        RowData[CurID] = value;
                    }
                }
            }

            public int Length
            {
                get
                {
                    return RowData.ItemArray.Length;
                }
            }

            #endregion 实现ENUM接口
        }

        /// <summary>
        /// 很多自定义的数据行，可以直接实现带着简繁体和去掉DBNull后的行。
        /// </summary>
        [Serializable]
        public class MyDataRowColl : IEnumerable<MyDataRow>, IEnumerator<MyDataRow>
        {
            #region 存储器
            private readonly DataRowCollection Rows;
            //private readonly DataRow[] Rows;
            private int CurID = -1;
            #endregion 存储器

            public DataRowCollection DataRows
            {
                get
                {
                    return Rows;
                }
            }

            /// <summary>
            /// 第一行或者唯一的一行。
            /// </summary>
            public MyDataRow FirstOrOnlyRow
            {
                get
                {
                    if (Rows != null && Rows.Count > 0)
                        return new MyDataRow(Rows[0]);
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
                    if (Rows != null)
                        return (Rows.Count > 0);
                    else
                        return false;
                }
            }

            #region 构造器

            public MyDataRowColl(DataRowCollection xRows)
            {
                Rows = xRows;
            }

            public MyDataRowColl(DataRow[] ArrayDataRow)
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
                        Rows = dt.Rows;
                        foreach (DataRow xaw in ArrayDataRow)
                        {
                            DataRow xxa = dt.NewRow();
                            foreach (DataColumn dCol in dc)
                            {
                                xxa[dCol.ColumnName] = xaw[dCol];
                            }
                            Rows.Add(xxa);
                        }
                    }
                }
            }

            public MyDataRowColl(DataTable xTable)
            {
                if (xTable != null)
                {
                    Rows = xTable.Rows;
                }
            }

            #endregion 构造器

            #region 索引器

            public MyDataRow this[int Index]
            {
                get
                {
                    CurID = Index;
                    return new MyDataRow(Rows[Index]);
                }
            }

            #endregion 索引器

            #region 实现枚举接口

            IEnumerator IEnumerable.GetEnumerator()
            {
                CurID = -1;
                return (IEnumerator)this;
            }

            public IEnumerator<MyDataRow> GetEnumerator()
            {
                CurID = -1;
                return (IEnumerator<MyDataRow>)this;
            }

            public void Reset()
            {
                CurID = -1;
            }

            public bool MoveNext()
            {
                CurID++;
                if (Rows != null && CurID >= 0 && CurID < Rows.Count)
                {
                    return true;
                }
                else
                {
                    CurID = -1;
                    return false;
                }
            }

            object IEnumerator.Current
            {
                get
                {
                    if (CurID >= 0 && CurID < Rows.Count)
                    {
                        return new MyDataRow(Rows[CurID]);
                    }
                    else
                    {
                        return null;
                    }
                }
            }

            public MyDataRow Current
            {
                get
                {
                    if (CurID >= 0 && CurID < Rows.Count)
                        return new MyDataRow(Rows[CurID]);
                    else
                        return null;
                }
            }

            public int Length
            {
                get
                {
                    return Rows == null ? -1 : Rows.Count;
                }
            }

            public int Count
            {
                get
                {
                    return Length;
                }
            }

            void IDisposable.Dispose()
            {
            }

            #endregion 实现枚举接口
        }

        /// <summary>
        /// 运行SQL语句命令。
        /// </summary>
        public class MyCommand
        {
            #region 存储器
            private MyCommandColl _myCmdColl;
            #endregion 存储器

            #region 运行无有返回值的SQL语句

            #region 各种重载，方便使用

            /// <summary>
            /// 运行一条SQL语句
            /// 注意：这将覆盖SQLCmdColl属性，将会使SQLCmdColl属性变成新的只有一条SQL的序列。
            /// 出错会回滚操作。
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <returns>会返回是否出错。</returns>
            public bool Execute(string SQL)
            {
                SQLCmdColl = new MyCommandColl();
                SQLCmdColl.Add(SQL);
                return Execute();
            }

            /// <summary>
            /// 运行一条SQL语句
            /// 注意：这将覆盖SQLCmdColl属性，将会使SQLCmdColl属性变成新的只有一条SQL的序列。
            /// 出错会回滚操作。可以很多参数
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="Params">很多参数，或者一个数组</param>
            /// <returns></returns>
            public bool Execute(string SQL, params MyParameter[] Params)
            {
                SQLCmdColl = new MyCommandColl();
                SQLCmdColl.Add(SQL, "Main", Params);
                return Execute();
            }

            /// <summary>
            /// 运行一条SQL语句
            /// 注意：这将覆盖SQLCmdColl属性，将会使SQLCmdColl属性变成新的只有一条SQL的序列。
            /// 出错会回滚操作。固定一个参数。
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="param0">固定一个参数</param>
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
            /// <param name="SQL">SQL语句</param>
            /// <param name="param0">第一个参数</param>
            /// <param name="param1">第二个参数</param>
            /// <returns></returns>
            public bool Execute(string SQL, MyParameter param0, MyParameter param1)
            {
                MyParameter[] ps = new MyParameter[2] { param0, param1 };
                return Execute(SQL, ps);
            }

            /// <summary>
            /// 运行一条SQL语句
            /// 注意：这将覆盖SQLCmdColl属性，将会使SQLCmdColl属性变成新的只有一条SQL的序列。
            /// 出错会回滚操作。固定三个参数
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="param0">第一个参数</param>
            /// <param name="param1">第二个参数</param>
            /// <param name="param2">第三个参数</param>
            /// <returns></returns>
            public bool Execute(string SQL, MyParameter param0, MyParameter param1, MyParameter param2)
            {
                MyParameter[] ps = new MyParameter[3] { param0, param1, param2 };
                return Execute(SQL, ps);
            }

            /// <summary>
            /// 运行SQLCmdColl属性中的SQL语句。
            /// 这些SQL语句是不返回值的。
            /// SQLList会覆盖SQLCmdColl属性。
            /// 运行语句是有事务的，出错会回滚操作。
            /// </summary>
            /// <param name="SQLList">设置要运行的语句序列，会覆盖SQLCmdColl属性</param>
            /// <returns>运行是否正确，不正确将RollBack</returns>
            public bool Execute(MyCommandColl SQLList)
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
                        using (SqlTransaction st = sn.BeginTransaction())
                        {
                            using (SqlCommand sc = new SqlCommand() { CommandType = CommandType.Text, Connection = sn, Transaction = st, CommandTimeout = 60000 })
                            {
                                try
                                {
                                    SQLCmdColl.Reset();
                                    int i = 0, l = SQLCmdColl.Count;
                                    foreach (MyCommandItem mc in SQLCmdColl)
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
                                            if (mc.Param != null)
                                            {
                                                sc.Parameters.Clear();
                                                sc.Parameters.AddRange(mc.Param);
                                            }
                                            mc.EffectRows = sc.ExecuteNonQuery(); //运行SQL语句。
                                            mc.ExecuteOK = true;
                                        }
                                        catch (Exception e)
                                        {
                                            MyRecord.Say(e, string.Format("SQLKey={0}", mc.key));
                                            mc.ExecuteOK = false;
                                            mc.ExecuteExp = e;
                                            throw e;
                                        }
                                    }
                                    st.Commit();
                                    ExecOK = true;
                                }
                                catch (Exception e)
                                {
                                    st.Rollback();
                                    throw e;
                                }
                                finally
                                {
                                    sc.Parameters.Clear();
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    MyRecord.Say(e, "", LogT);
                    ExecOK = false;
                }

                return ExecOK;
            }

            #endregion 运行无有返回值的SQL语句

            #region 多条SQL语句

            /// <summary>
            /// SQL语句序列
            /// </summary>
            public MyCommandColl SQLCmdColl
            {
                get
                {
                    if (_myCmdColl == null)
                    {
                        _myCmdColl = new MyCommandColl();
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
            /// <param name="SQL">SQL语句</param>
            /// <param name="Key">KEY</param>
            /// <param name="args">参数序列</param>
            public void Add(string SQL, string Key, params MyParameter[] args)
            {
                SQLCmdColl.Add(SQL, Key, args);
            }

            /// <summary>
            /// 加一个SQL语句
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="Key">KEY</param>
            public void Add(string SQL, string Key)
            {
                SQLCmdColl.Add(SQL, Key);
            }

            /// <summary>
            /// 加入一个SQL语句，没有KEY，出错时将不知道哪里错误。
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            public void Add(string SQL)
            {
                SQLCmdColl.Add(SQL);
            }

            /// <summary>
            /// 加入一个SQL语句对象
            /// </summary>
            /// <param name="SQLItem">SQL语句对象</param>
            public void Add(MyCommandItem SQLItem)
            {
                SQLCmdColl.Add(SQLItem);
            }

            #endregion 加载SQL语句

            #endregion 多条SQL语句

            #region 多条SQL语句存储类

            /// <summary>
            /// 要运行的一条SQL命令。
            /// </summary>
            public class MyCommandItem
            {
                #region 存储器

                public string SQL { get; set; }

                public System.Exception ExecuteExp { get; set; }

                public bool ExecuteOK { get; set; }

                public string key { get; set; }

                public int index { get; set; }

                public int EffectRows { get; set; }

                public SqlParameter[] Param { get; set; }

                public MyParameter[] myParam { get; set; }

                #endregion 存储器

                #region 构造器

                public MyCommandItem()
                {
                    index = -1;
                }

                public MyCommandItem(string sql)
                {
                    SQL = sql;
                    key = string.Empty;
                    index = -1;
                }

                public MyCommandItem(string sql, string Key)
                {
                    SQL = sql;
                    key = Key;
                    index = -1;
                }

                public MyCommandItem(string sql, string Key, params MyParameter[] args)
                {
                    SQL = sql;
                    key = Key;
                    index = -1;
                    myParam = args;
                    Param = new SqlParameter[args.Length];
                    for (int i = 0; i < args.Length; i++)
                    {
                        Param[i] = args[i].sqlparam;
                    }
                }

                #endregion 构造器
            }

            /// <summary>
            /// 保持要运行的SQL集合。
            /// </summary>
            public class MyCommandColl : IEnumerator, IEnumerable
            {
                #region 存储器
                private readonly ArrayList IDPointer;
                private int Cid = -1;
                private readonly Dictionary<string, MyCommandItem> MyColl;
                #endregion 存储器

                #region 构造器

                public MyCommandColl()
                {
                    MyColl = new Dictionary<string, MyCommandItem>();
                    IDPointer = new ArrayList();
                }

                #endregion 构造器

                #region 加减元素

                public void Add(MyCommandItem SQLCmdItem)
                {
                    MyColl.Add(SQLCmdItem.key, SQLCmdItem);
                    IDPointer.Add(SQLCmdItem.key);
                    Cid++;
                    SQLCmdItem.index = Cid;
                }

                public void Add(string SQL, string KEY, params MyParameter[] args)
                {
                    Add(new MyCommandItem(SQL, KEY, args));
                }

                public void Add(string SQL, string KEY)
                {
                    Add(new MyCommandItem(SQL, KEY));
                }

                public void Add(string SQL)
                {
                    Add(new MyCommandItem(SQL));
                }

                public int Count
                {
                    get
                    {
                        return MyColl.Count;
                    }
                }

                public void Remove(string KEY)
                {
                    MyColl.Remove(KEY);
                    IDPointer.Remove(KEY);
                }

                #endregion 加减元素

                #region 实现枚举接口

                public IEnumerator GetEnumerator()
                {
                    return (IEnumerator)this;
                }

                public void Reset()
                {
                    Cid = -1;
                }

                public bool MoveNext()
                {
                    Cid++;
                    return Cid < MyColl.Count;
                }

                public object Current
                {
                    get
                    {
                        return MyColl[(string)IDPointer[Cid]];
                    }
                    set
                    {
                        MyColl[(string)IDPointer[Cid]] = (MyCommandItem)value;
                    }
                }

                #endregion 实现枚举接口
            }

            #endregion 多条SQL语句存储类

            #region 运行一个返回一个值的一条SQL语句

            /// <summary>
            /// 从数据库返回一个值
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="paramArgs">参数序列</param>
            /// <returns>一个值，注意要（类型）objec强制类型转换</returns>
            public static object ExecuteScalar(string SQL, params MyData.MyParameter[] paramArgs)
            {
                Type LogT = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType;
                try
                {
                    SqlConnection cnn = new SqlConnection(MyBase.ConnectionString);
                    cnn.Open();
                    SQL =MyConvert.LCDB(SQL).ToString(); //进行繁简转换。
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
                        _R = _R == DBNull.Value ? null : _R;
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
            /// <param name="SQL">SQL语句</param>
            /// <returns>一个值，注意要（类型）objec强制类型转换</returns>
            public static object ExecuteScalar(string SQL)
            {
                return ExecuteScalar(SQL, (MyData.MyParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值，支持{0}参数，内部用string.Format自动格式化。
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="strArgs">字符串参数序列</param>
            /// <returns>一个值，注意要（类型）objec强制类型转换</returns>
            public static object ExecuteScalar(string SQL, params string[] strArgs)
            {
                return ExecuteScalar(string.Format(SQL, strArgs), (MyData.MyParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值。
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="strArg0">第一个参数</param>
            /// <returns>一个值，注意要（类型）objec强制类型转换</returns>
            public static object ExecuteScalar(string SQL, string strArg0)
            {
                return ExecuteScalar(string.Format(SQL, strArg0), (MyData.MyParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值。
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="strArg0">第一个参数</param>
            /// <param name="strArg1">第二个参数</param>
            /// <returns></returns>
            public static object ExecuteScalar(string SQL, string strArg0, string strArg1)
            {
                return ExecuteScalar(string.Format(SQL, strArg0, strArg1), (MyData.MyParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="strArg0">第一个参数</param>
            /// <param name="strArg1">第二个参数</param>
            /// <param name="strArg2">第三个参数</param>
            /// <returns></returns>
            public static object ExecuteScalar(string SQL, string strArg0, string strArg1, string strArg2)
            {
                return ExecuteScalar(string.Format(SQL, strArg0, strArg1, strArg2), (MyData.MyParameter[])null);
            }

            /// <summary>
            /// 从数据库返回一个值,用SQL参数的。
            /// </summary>
            /// <param name="SQL">SQL语句</param>
            /// <param name="paramArg0">第一个参数</param>
            /// <returns></returns>
            public static object ExecuteScalar(string SQL, MyData.MyParameter paramArg0)
            {
                MyData.MyParameter[] mp = new MyParameter[1];
                mp[0] = paramArg0;
                return ExecuteScalar(SQL, mp);
            }

            #endregion 运行一个返回一个值的一条SQL语句
        }

        /// <summary>
        /// SQL语句参数类。
        /// </summary>
        [Serializable]
        public class MyParameter
        {
            [Serializable]
            public enum MyDataType : uint
            {
                Boolean,
                DateTime,
                Int,
                Numeric,
                String
            }

            private object _Value = null;
            private string _Name = string.Empty;
            private MyDataType _MyDbType = MyDataType.String;
            public SqlParameter sqlparam;

            public MyParameter()
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
            /// <param name="name"></param>
            /// <param name="value"></param>
            public MyParameter(string name, object value)
            {
                sqlparam = new SqlParameter();
                MyDbType = MyDataType.String;
                Value = value;
                Name = name;

                //sqlparam = new SqlParameter(_Name, _Value);
            }

            public MyParameter(string name, object value, MyDataType mtype)
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
                                                _Value = Convert.ToDouble(value);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (_MyDbType == MyDataType.Boolean)
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
                            else
                            {
                                if (_MyDbType == MyDataType.DateTime)
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


        #endregion 数据返回的固定信息。
    }
}