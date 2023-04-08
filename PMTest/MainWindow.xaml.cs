using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerMILL;
using System.Text.RegularExpressions;

namespace PMTest
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        Dictionary<int, PowerMILL.Application> pmDic = new Dictionary<int, PowerMILL.Application>();//主数据列表 使用字典便于管理

        PowerMILL.Application ActivePm = null;//当前控制激活的实例


        private async void Start_Button_Click(object sender, RoutedEventArgs e)
        {

            while (true)//循环开始
            {

                //创建一个临时字典数据列表
                Dictionary<int, PowerMILL.Application> pmDic_ = new Dictionary<int, PowerMILL.Application>();

                //获取所有已打开的PowerMill实例并放入临时字典数据列表
                foreach (var item in GetRunningInstances("PowerMILL.Application"))
                {
                    PowerMILL.Application PMA_ = (PowerMILL.Application)item;
                    pmDic_.Add(PMA_.ParentWindow, PMA_);
                }

                //检查pmDic列表多余的数据项准备删除
                var RemoveDic = pmDic.Except(pmDic_);

                //检查pmDic列表没有的数据项准备添加
                var Addindic = pmDic_.Except(pmDic);

                //删除项循环
                foreach (var item in RemoveDic)
                {
                    //删除对应的Button控件
                    for (int i = 0; i < PMCortol.Children.Count; i++)
                    {
                        if (((Button)PMCortol.Children[i]).Tag.ToString() == item.Key.ToString()) PMCortol.Children.RemoveAt(i);
                    }
                }
                pmDic.Except(RemoveDic);//删除数据

                //添加项循环
                foreach (var item in Addindic)
                {
                    pmDic.Add(item.Key, item.Value);//添加到字段数据
                    Button button = new Button();//新建Button
                    button.Content = item.Key.ToString();//按钮显示文本
                    button.Tag = item.Key.ToString();//Tag
                    button.Background = Brushes.Transparent;//背景
                    button.Margin = new Thickness(5, 5, 0, 0);//外边距
                    button.Cursor = Cursors.Hand;//鼠标形状
                    button.Click += ActiveObject;//点击事件
                    PMCortol.Children.Add(button);//Button添加到StackPanel
                }

                //识别Powermill打开的项目名称输入到Button显示文本上
                foreach (Button item in PMCortol.Children)
                {
                    string Value_ = PsGetVal_(pmDic[(int.Parse(item.Tag.ToString()))], "project_pathname(1)");
                    if (Value_ != null) item.Content = Value_;
                }

                //线程暂停5秒
                await Task.Delay(5000);
            }
            
        }


        private void ActiveObject(object sender, RoutedEventArgs e)
        {
            Button button = (Button)sender;
            ActivePm = pmDic[(int.Parse(button.Tag.ToString()))];
            foreach (Button item in PMCortol.Children)
            {
                if (item== button)
                {
                    item.Background = Brushes.Teal;
                    
                }
                else
                {
                    item.Background = Brushes.Transparent;
                }
                
            }
        }



        #region GetRunningObjectTable=>获取所有Powermill实例
        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, ref IBindCtx ppbc);

        [DllImport("ole32.dll")]
        public static extern void GetRunningObjectTable(int reserved, ref IRunningObjectTable prot);
        public static List<object> GetRunningInstances(string classId)
        {
            List<string> list = new List<string>();
            Type typeFromProgID = Type.GetTypeFromProgID(classId);
            list.Add(typeFromProgID.GUID.ToString().ToUpper());
            IRunningObjectTable prot = null;
            GetRunningObjectTable(0, ref prot);
            if (prot == null)
            {
                return null;
            }

            IEnumMoniker ppenumMoniker = null;
            prot.EnumRunning(out ppenumMoniker);
            if (ppenumMoniker == null)
            {
                return null;
            }

            ppenumMoniker.Reset();
            List<object> list2 = new List<object>();
            IntPtr pceltFetched = default(IntPtr);
            IMoniker[] array = new IMoniker[1];
            List<string> list3 = new List<string>();
            while (ppenumMoniker.Next(1, array, pceltFetched) == 0)
            {
                IBindCtx ppbc = null;
                CreateBindCtx(0u, ref ppbc);
                if (ppbc == null)
                {
                    continue;
                }

                string ppszDisplayName = null;
                array[0].GetDisplayName(ppbc, null, out ppszDisplayName);
                list3.Add(ppszDisplayName);
                foreach (string item in list)
                {
                    _ = item;
                    if (ppszDisplayName.Contains(classId))
                    {
                        object ppunkObject = null;
                        prot.GetObject(array[0], out ppunkObject);
                        if (ppunkObject != null)
                        {
                            list2.Add(ppunkObject);
                            break;
                        }
                    }
                }
            }
            return list2;
        }

        #endregion

        #region PmCom
        public void Com(string comd)
        {
            //ps.DoCommand(token, comd);
            ActivePm.DoCommand( comd);
        }
        private string ComEx(string comd)
        {
            object item;
            ActivePm.DoCommand( "ECHO OFF DCPDEBUG UNTRACE COMMAND ACCEPT");
            ActivePm.DoCommandEx(comd, out item);
            return item.ToString().TrimEnd();
        }
        public string PsGetVal(string comd)
        {
            Regex rg = new Regex("(?<=(" + ">" + "))[.\\s\\S]*?(?=(" + "<" + "))", RegexOptions.Multiline | RegexOptions.Singleline);
            return rg.Match(ActivePm.GetParameterXML(comd)).Value;
        }
        public string PsGetVal_(PowerMILL.Application pm_,string str)
        {
            string returnstr = null;
            if (!pm_.Busy)
            {
                Regex rg = new Regex("(?<=(" + ">" + "))[.\\s\\S]*?(?=(" + "<" + "))", RegexOptions.Multiline | RegexOptions.Singleline);
                returnstr= rg.Match(pm_.GetParameterXML(str)).Value;
            }
            return returnstr;
        }
        #endregion

        private void Show_Test(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ActivePm != null)
                {
                    Com($"print='当前控制了这个窗口:{DateTime.Now.ToString()}'");
                }

            }
            catch (Exception ex)
            {
               Console.WriteLine("测试出现了一些错误\r"+ex.ToString());
            }
        }
    }
}
