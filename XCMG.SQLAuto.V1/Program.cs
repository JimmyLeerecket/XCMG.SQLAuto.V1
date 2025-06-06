using XCMG.SQLAuto.V1;

public partial class Program
{
    static void Main(string[] args)
    {
        // 生成统计单
        string filePath = "C:\\Mac\\Home\\Desktop\\jimmyli\\Import\\Input\\SQL导入模版.xlsx";
        Helper.ImportExcel(filePath);
    }
}

