using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using System.Xml;
using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;


namespace MYERP_ServerServiceRuner.Base
{
    static public class MyBase
    {
        public static string ConnectionString
        {
            get
            {
                return string.Format("Data Source={0};Persist Security Info=True;Password={2};User ID={1};Initial Catalog={3};Application Name=ERPSVR_{5}_{4};Connection Timeout=300",
                    ServerAddress, DB_UID, DB_PWD, DB_Name, System.Environment.MachineName, Application.ProductVersion);
            }
        }

        public static string ServerAddress
        {
            get
            {
                return ConfigurationManager.AppSettings["Server"];
            }
        }

        public static string DB_UID
        {
            get
            {
                return ConfigurationManager.AppSettings["UID"];
            }
        }

        public static string DB_PWD
        {
            get
            {
                return ConfigurationManager.AppSettings["PWD"];
            }
        }

        public static string DB_Name
        {
            get
            {
                return ConfigurationManager.AppSettings["Database"];
            }
        }

        public static string CompanyTitle
        {
            get
            {
                return ConfigurationManager.AppSettings["Title"];
            }
        }

        /// <summary>
        /// 是空记录集
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="o"></param>
        /// <returns></returns>
        public static bool IsEmptySet<T>(this IEnumerable<T> o)
        {
            if (o.IsNotNull())
            {
                return o.Count() <= 0;
            }
            return true;
        }
        /// <summary>
        /// 不是空记录集。
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="o"></param>
        /// <returns></returns>
        public static bool IsNotEmptySet<T>(this IEnumerable<T> o)
        {
            if (o.IsNotNull())
            {
                return o.Count() > 0;
            }
            return false;
        }



        #region 集合类——泛型基类

        #region 集合类——泛型基类
        /// <summary>
        /// 必须重写索引器this[Key]
        /// </summary>
        /// <typeparam Name="T"></typeparam>
        [Serializable]
        public abstract class MyEnumerable<T> : IEnumerable<T>, IEnumerator<T>, IDisposable
        {
            #region 构造函数和析构函数

            #endregion 构造函数和析构函数

            #region 存储器

            protected List<T> _data = new List<T>();

            #endregion 存储器

            #region 索引器

            public virtual T this[int index]
            {
                get
                {
                    return _data[index];
                }
                set
                {
                    _data[index] = value;
                }
            }

            public virtual T this[Func<T, bool> FindAction]
            {
                get
                {
                    var v = from a in _data
                            where FindAction(a)
                            select a;
                    return v.FirstOrDefault();
                }
                set
                {
                    var v = from a in _data
                            where FindAction(a)
                            select a;
                    T x = v.FirstOrDefault();
                    x = value;
                }
            }

            #endregion 索引器

            #region 实现枚举接口

            protected int _current_index = -1;

            object IEnumerator.Current
            {
                get
                {
                    if (_current_index >= 0 && _current_index < _data.Count)
                        return this[_current_index];
                    else
                        return null;
                }
            }

            public virtual T Current
            {
                get
                {
                    if (_current_index >= 0 && _current_index < _data.Count)
                        return this[_current_index];
                    else
                        return default(T);
                }
                set
                {
                    this[_current_index] = (T)value;
                }
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                _current_index = -1;
                return (IEnumerator)this;
            }

            public IEnumerator<T> GetEnumerator()
            {
                _current_index = -1;
                return (IEnumerator<T>)this;
            }

            public virtual int Length
            {
                get
                {
                    if (_data.IsNotNull()) return _data.Count; else return 0;
                }
            }

            public virtual int Count()
            {
                if (_data.IsNotNull()) return _data.Count; else return 0;
            }

            public bool MoveNext()
            {
                _current_index++;
                return (_data != null && _current_index < _data.Count);
            }

            public void Reset()
            {
                _current_index = -1;
            }

            void IDisposable.Dispose()
            {
                //_data = null;
            }

            #endregion 实现枚举接口

            #region 操作函数

            /// <summary>
            /// 增加一个元素
            /// </summary>
            /// <param Name="item">元素</param>
            public virtual void Add(T item)
            {
                if (_data.IsNull()) _data = new List<T>();
                _data.Add(item);
            }

            /// <summary>
            /// 增加一个元素
            /// </summary>
            /// <returns>返回增加的元素是空值或者0</returns>
            public virtual T Add()
            {
                T xItem = default(T);
                Add(xItem);
                return xItem;
            }

            /// <summary>
            /// 增加很多元素
            /// </summary>
            /// <param Name="proditems"></param>
            public virtual void AddRange(T[] ArrayItems)
            {
                if (ArrayItems != null && ArrayItems.Length > 0)
                {
                    foreach (T xItem in ArrayItems)
                    {
                        Add(xItem);
                    }
                }
            }

            /// <summary>
            /// 清除所有元素
            /// </summary>
            public virtual void Clear()
            {
                _data.Clear();
                _current_index = -1;
            }

            /// <summary>
            /// 去掉一个元素
            /// </summary>
            /// <param Name="item">要去掉的元素</param>
            public virtual void RemoveItem(T item)
            {
                _data.Remove(item);
            }

            public bool Contains(Func<T, bool> FindAction)
            {
                var v = from a in _data
                        where FindAction(a)
                        select a;
                return v.IsNotEmptySet();
            }

            #endregion 操作函数
        }

        #endregion 集合类——泛型基类

        #endregion 集合类——泛型基类


        public class SendMail
        {
            public MailAddress MailFrom
            {
                get
                {
                    return (new MailAddress("AutoMessage@my.imedia.com.tw", "ERP[伺服程式]", Encoding.UTF8));
                }
            }

            public string MailBodyText { get; set; }

            public string MailTo { get; set; }

            public string MailCC { get; set; }

            public string Subject { get; set; }

            public SendMail()
            {
                mail = new MailMessage();
                Attachments = mail.Attachments;
            }

            public MailMessage mail { get; set; }

            public AttachmentCollection Attachments { get; set; }

            public void SendOut()
            {
                mail.From = MailFrom;
                mail.Subject = MyConvert.ZH_TW(Subject);
                mail.Sender = MailFrom;
                mail.SubjectEncoding = Encoding.UTF8;
                mail.IsBodyHtml = true;
                mail.Body = MyConvert.ZH_TW(MailBodyText);
                mail.To.Add(MailTo);
                if (MailCC != null && MailCC.Length > 0) mail.CC.Add(MailCC);
                mail.BodyEncoding = Encoding.UTF8;
                mail.HeadersEncoding = Encoding.UTF8;
                mail.Headers.Add("SenderIPAddress", LocalInfo.GetLocalIp());
                mail.Headers.Add("Sender", MyConvert.ZHLC("ERP自动化伺服器"));
                mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure | DeliveryNotificationOptions.Delay;
                mail.Priority = MailPriority.Normal;

                //SmtpClient Sender = new SmtpClient("59.125.179.177");
                SmtpClient Sender = new SmtpClient("my.imedia.com.tw");
                Sender.DeliveryMethod = SmtpDeliveryMethod.Network;
                Sender.EnableSsl = false;
                Sender.Port = 25;
                Sender.UseDefaultCredentials = false;
                Sender.Timeout = 120 * 1000;
                Sender.Credentials = new NetworkCredential("AutoMessage", "HWGV1z86");

                try
                {
                    Sender.Send(mail);
                }
                catch (Exception ex)
                {
                    MyRecord.Say("再试一次");
                    MyRecord.Say(ex);
                    try
                    {
                        Sender.Send(mail);
                    }
                    catch (Exception exx)
                    {
                        MyRecord.Say(exx);
                    }
                }
            }


        }

    }


    public static class LocalInfo
    {
        #region 获取本机IP

        /// <summary>
        /// 获取本机IP
        /// </summary>
        /// <returns></returns>
        public static string GetLocalIp()
        {
            try
            {
                string hostname = Dns.GetHostName();//得到本机名
                IPHostEntry localhost = Dns.GetHostEntry(hostname);
                IPAddress localaddr = localhost.AddressList[0];
                foreach (IPAddress ipItem in localhost.AddressList)
                {
                    if (ipItem.IsIPv6SiteLocal || ipItem.IsIPv6Multicast || ipItem.IsIPv6LinkLocal || ipItem.IsIPv6Teredo)
                    {
                        continue;
                    }
                    if (ipItem != null)
                    {
                        Regex rg = new Regex(@"^(192\.168\.\d{1,2}\.)(\d{1,2}|1\d\d|2[0-4]\d|25[0-5])$");
                        if (rg.IsMatch(ipItem.ToString()))
                        {
                            localaddr = ipItem;
                            break;
                        }
                        else
                            continue;
                    }
                }
                return localaddr.ToString();
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion 获取本机IP
    }

    public static class MyConvert
    {
        #region 获取操作系统的语言类型操作。

        /// <summary>
        /// 得到操作系统的语言类别的API。
        /// </summary>
        /// <returns></returns>
        [DllImport("kernel32", EntryPoint = "GetSystemDefaultLCID", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int GetSystemDefaultLCID();

        /// <summary>
        /// 操作系统语言类别
        /// </summary>
        public enum MY_OSType
        {
            ChineseSimplified,
            ChineseTradiational,
            English,
            Japanese,
            Other
        }

        /// <summary>
        /// 得到操作系统的语言类别。
        /// </summary>
        public static MY_OSType GetOSType
        {
            get
            {
                int i = GetSystemDefaultLCID();
                switch (i)
                {
                    case 2052:
                        return MY_OSType.ChineseSimplified;
                    case 1028:
                        return MY_OSType.ChineseTradiational;
                    case 1033:
                        return MY_OSType.English;
                    case 1041:
                        return MY_OSType.Japanese;
                    default:
                        return MY_OSType.Other;
                }
            }
        }

        #endregion 获取操作系统的语言类型操作。
        /// <summary>
        /// 简体到繁体
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string ZH_TW(object s)
        {
            if (s == null)
            {
                return string.Empty;
            }
            if (s is DBNull)
            {
                return string.Empty;
            }
            return ChineseConverter.Convert(s.ToString(), ChineseConversionDirection.SimplifiedToTraditional);
        }

        /// <summary>
        /// 繁体到简体
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string TW_ZH(object s)
        {
            if (s == null) s = string.Empty;
            return ChineseConverter.Convert(s.ToString(), ChineseConversionDirection.TraditionalToSimplified);
        }

        /// <summary>
        /// 数据库到本地（数据库是繁体)
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string DBLC(object str)
        {
            MY_OSType myosType = GetOSType;
            if (str == null) str = string.Empty;
            if (myosType == MY_OSType.ChineseSimplified)
            {
                return TW_ZH(str);
            }
            else
            {
                string s = str.ToString();
                if (myosType == MY_OSType.ChineseTradiational)
                {
                    return ZH_TW(s);
                }
                else
                {
                    if (myosType == MY_OSType.English)
                    {
                        return s;
                    }
                    else
                    {
                        if (myosType == MY_OSType.Japanese)
                        {
                            return s;
                        }
                        else
                        {
                            if (myosType == MY_OSType.Other)
                            {
                                return s;
                            }
                            else
                            {
                                return s;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 本地到数据库，（数据库是繁体）
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static object LCDB(object s)
        {
            MY_OSType myosType = GetOSType;
            if (myosType == MY_OSType.ChineseSimplified)
            {
                return ZH_TW(s);
            }
            else
            {
                if (myosType == MY_OSType.ChineseTradiational)
                {
                    return ZH_TW(s);
                }
                else
                {
                    if (myosType == MY_OSType.English)
                    {
                        return s;
                    }
                    else
                    {
                        if (myosType == MY_OSType.Japanese)
                        {
                            return s;
                        }
                        else
                        {
                            if (myosType == MY_OSType.Other)
                            {
                                return s;
                            }
                            else
                            {
                                return s;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 简体到本地，（程序内部输入的汉字必须要转换）
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string ZHLC(object str)
        {
            MY_OSType myosType = GetOSType;
            if (str == null) str = string.Empty;
            string s = str.ToString();
            if (myosType == MY_OSType.ChineseSimplified)
            {
                return TW_ZH(s);
            }
            else
            {
                if (myosType == MY_OSType.ChineseTradiational)
                {
                    return ZH_TW(s);
                }
                else
                {
                    if (myosType == MY_OSType.English)
                    {
                        return s;
                    }
                    else
                    {
                        if (myosType == MY_OSType.Japanese)
                        {
                            return s;
                        }
                        else
                        {
                            if (myosType == MY_OSType.Other)
                            {
                                return s;
                            }
                            else
                            {
                                return s;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 本地到简体，用于自写配置转回。
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static string LCZH(object str)
        {
            MY_OSType myosType = GetOSType;
            if (str == null) str = string.Empty;
            string s = str.ToString();
            if (myosType == MY_OSType.ChineseSimplified)
            {
                return TW_ZH(s);
            }
            else
            {
                if (myosType == MY_OSType.ChineseTradiational)
                {
                    return TW_ZH(s);
                }
                else
                {
                    if (myosType == MY_OSType.English)
                    {
                        return s;
                    }
                    else
                    {
                        if (myosType == MY_OSType.Japanese)
                        {
                            return s;
                        }
                        else
                        {
                            if (myosType == MY_OSType.Other)
                            {
                                return s;
                            }
                            else
                            {
                                return s;
                            }
                        }
                    }
                }
            }
        }
    }

    public static class MyRecord
    {
        private static log4net.ILog _logger = log4net.LogManager.GetLogger("");

        public static log4net.ILog Logger
        {
            get
            {
                return _logger;
            }
        }

        public static void Say(Exception E)
        {
            Say(E, string.Empty);
        }

        /// <summary>
        /// 说出错误
        /// </summary>
        /// <param name="E"></param>
        /// <param name="DesWord"></param>
        public static void Say(Exception E, string DesWord)
        {
            string errorMessage = MyConvert.ZHLC(string.Format("\n{0}", DesWord));
            _logger.Error(errorMessage, E);
        }

        public static void Say(Exception E, string DesWord, Type type)
        {
            log4net.ILog _lclogger = log4net.LogManager.GetLogger(type);
            string errorMessage = MyConvert.ZHLC(string.Format("\n{0}", DesWord));
            _lclogger.Error(errorMessage, E);
        }

        public static void Say(string s)
        {
            string inf = MyConvert.ZHLC(string.Format("{0}", s));
            _logger.Info(inf);
        }

        public static void Say(string s, Type type)
        {
            log4net.ILog _lclogger = log4net.LogManager.GetLogger(type);
            string inf = MyConvert.ZHLC(string.Format("{0}", s));
            _lclogger.Info(inf);
        }
    }

}
