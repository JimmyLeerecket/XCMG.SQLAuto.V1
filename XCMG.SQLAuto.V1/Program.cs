using System.Text.RegularExpressions;
using XCMG.SQLAuto.V1;
using XCMG.SQLAuto.V1.Study;

public partial class Program
{
    static void Main(string[] args)
    {
        // 生成统计单
        string filePath = "C:\\Mac\\Home\\Desktop\\jimmyli\\Import\\Input\\SQL导入模版_DL.xlsx";
        Helper.ImportExcel(filePath);

        //var currentThread = Thread.CurrentThread;
        //Console.WriteLine("线程标识：" + currentThread.Name);
        //Console.WriteLine("当前地域：" + currentThread.CurrentCulture.Name);
        //Console.WriteLine("线程执行状态：" + currentThread.IsAlive);
        //Console.WriteLine("是否为后台线程：" + currentThread.IsBackground);
        //Console.WriteLine("是否为线程池线程" + currentThread.IsThreadPoolThread);

        // LINQ
        // LINQ.SelectNewObject();
    }
}

