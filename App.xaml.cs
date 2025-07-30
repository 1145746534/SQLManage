using Prism.Ioc;
using Prism.Unity;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Unity;
using Prism;
using SQLManage.Views;
using System.Runtime.InteropServices;
using MahApps.Metro.Controls.Dialogs;

namespace SQLManage
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : PrismApplication
    {
        // 用于进程间通信的Mutex名称（需唯一）
        private string MutexName = "baobiaoguanli12574";

        private static Mutex _appMutex;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        //根据任务栏应用程序显示的名称找窗口的名称
        [DllImport("User32.dll", EntryPoint = "FindWindow")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        private const int SW_RESTORE = 9;
        private const int SW_MAXIMIZE = 3;

        protected override void OnStartup(StartupEventArgs e)
        {

            try
            {
                // 尝试创建 Mutex
                _appMutex = new Mutex(true, MutexName, out bool createdNew);

                if (!createdNew)
                {
                    IntPtr findPtr = FindWindow(null, "报表管理");
                    if (findPtr.ToInt32() != 0)
                    {
                        ShowWindow(findPtr, SW_RESTORE); //将窗口还原，如果不用此方法，缩小的窗口不能激活
                        SetForegroundWindow(findPtr);//将指定的窗口选中(激活)
                    }
                    //ActivateExistingInstance();
                    Current.Shutdown();
                    return;
                }
            }
            catch (Exception ex)
            {
                // 记录错误但继续启动
                Debug.WriteLine($"Mutex 创建失败: {ex.Message}");
            }

            base.OnStartup(e);

        }

        protected override Window CreateShell()
        {
            return Container.Resolve<Views.MainWindow>();
            //throw new NotImplementedException();
        }
        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.RegisterForNavigation<ReportManagementView>();

            containerRegistry.RegisterInstance<IDialogCoordinator>(DialogCoordinator.Instance);

        }

    }
}
