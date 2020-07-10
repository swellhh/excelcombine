using System;

namespace excelmultipletosingle
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("解析开始");

            manage.fillconfigattr();
            manage.ExportOrgExcel();

            Console.WriteLine("解析结束");
            Console.ReadKey();

        }
    }
}
