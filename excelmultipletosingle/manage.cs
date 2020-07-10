using excelmultipletosingle.entity;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using ServiceStack.Text;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace excelmultipletosingle
{
    public static class manage
    {
        public static List<config> configs;


        public static void fillconfigattr()
        {
            var text = File.ReadAllText("C://Users//78575//Desktop//江河湖库数据处理//excelmultipletosingle//excelmultipletosingle//config.json");

            var configlist = JsonSerializer.DeserializeFromString<List<config>>(text);

            configs = configlist;


        }

        private static Dictionary<string,Dictionary<string, List<string>>> ReadDataFromExcel()
        {
            IWorkbook workbook = null;
            ISheet sheet = null;
            var res = new Dictionary<string, Dictionary<string, List<string>>>();

            foreach (var config in configs)
            {
                var resDic = new Dictionary<string, List<string>>();

                if (File.Exists(config.path))
                {
                    var file = new FileInfo(config.path);

                    FileStream fsw = new FileStream(config.path, FileMode.OpenOrCreate);

                    //判断excel版本
                    if (file.Extension.EndsWith(".xls"))
                    {
                        workbook = new HSSFWorkbook(fsw);
                    }
                    else if (file.Extension.EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(fsw);
                    }

                    fsw.Close();
                    fsw.Dispose();

                    sheet = workbook.GetSheet(config.sheetName);

                    if (sheet == null)
                    {
                        throw new Exception("sheet页不存在");
                    }

                    //获取总行数
                    int rowCount = sheet.PhysicalNumberOfRows;
                    //大于两行代表有数据
                    if (rowCount > 2)
                    {
                        for (var fi=0;fi<config.fileds.Length;fi++)
                        {
                            resDic.Add(config.fileds[fi], null);
                            var fieldValue = new List<string>();

                            for (int i = 0; i < rowCount; i++)
                            {
                                var row = sheet.GetRow(i + 1);
                                if (row != null)
                                {

                                    for (var j = 0; j < config.lastcol / config.fileds.Length; j++)
                                    {

                                        var space = config.space * j;
                                        var cell = row.GetCell(j * config.fileds.Length + fi + space);

                                        if (cell!=null)
                                        {
                                            if (config.fileds[fi] == "TM")
                                            {
                                                fieldValue.Add(cell.DateCellValue.ToString("yyyy-MM-dd HH:mm:ss"));

                                            }
                                            else
                                            {
                                                fieldValue.Add(cell.ToString());

                                            }
                                        }

                                    }
                                }
                            }

                            resDic[config.fileds[fi]] = fieldValue;
                        }


                    }
                }

                res.Add(config.path,resDic);


            }

            return res;
        }

        public static void ExportOrgExcel()
        {
            try
            {
                var data = ReadDataFromExcel();
                // 下面为导出文件
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("导入数据");
                IRow row0 = sheet.CreateRow(0);

                // 创建头部
                row0.CreateCell(0).SetCellValue("STCD");
                row0.CreateCell(1).SetCellValue("STNM");
                row0.CreateCell(2).SetCellValue("TM");
                row0.CreateCell(3).SetCellValue("DRP");

                foreach (var config in configs)
                {
                    
                    var overmaxrow = data[config.path][config.fileds[0]].Count / 1048575;

                    if (overmaxrow > 0)
                    {
                        for (var tuple = 1; tuple <= overmaxrow + 1; tuple++)
                        {
                            var sheet2 = workbook.CreateSheet("剩余导入数据" + tuple);

                            if (tuple == overmaxrow + 1)
                            {
                                for (int i = 0; i < data[config.path][config.fileds[0]].Count % 1048575; i++)
                                {
                                    IRow row = sheet2.CreateRow(i);

                                    for (var fi = 0; fi < config.fileds.Length; fi++)
                                    {
                                        ICell cell = row.CreateCell(fi);


                                        IDataFormat dataFormatCustom = workbook.CreateDataFormat();
                                        cell.CellStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");

                                        cell.SetCellValue(data[config.path][config.fileds[fi]][i + (tuple-1) * 1048575]);
                                    }


                                }

                            }

                            else
                            {
                                for (int i = 0; i < 1048575; i++)
                                {
                                    IRow row = sheet2.CreateRow(i + 1);

                                    for (var fi = 0; fi < config.fileds.Length; fi++)
                                    {
                                        ICell cell = row.CreateCell(fi);

                                        IDataFormat dataFormatCustom = workbook.CreateDataFormat();
                                        cell.CellStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");

                                        cell.SetCellValue(data[config.path][config.fileds[fi]][(tuple-1) * 1048575 + i]);
                                    }
                                }
                            }

                        }
                        
                    }

                    else
                    {
                        for (int i = 0; i < data[config.path][config.fileds[0]].Count; i++)
                        {
                            IRow row = sheet.CreateRow(i + 1);

                            for (var fi = 0; fi < config.fileds.Length; fi++)
                            {
                                ICell cell = row.CreateCell(fi);

                                IDataFormat dataFormatCustom = workbook.CreateDataFormat();
                                cell.CellStyle.DataFormat = dataFormatCustom.GetFormat("yyyy-MM-dd HH:mm:ss");

                                cell.SetCellValue(data[config.path][config.fileds[fi]][i]);
                            }
                        }
                    }

                    var file = new FileInfo(config.path);

                    Directory.CreateDirectory(Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory + "/Files/"));

                    var returnname = "/Files/"+ file.Name.Replace(file.Extension,"")+ "入库.xlsx";

                    using (FileStream url = File.OpenWrite(AppDomain.CurrentDomain.BaseDirectory + returnname))
                    {
                        // 导出Excel文件
                        workbook.Write(url);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("解析出错\n");
                Console.WriteLine(ex.Message);

            }

        }
    }
}
