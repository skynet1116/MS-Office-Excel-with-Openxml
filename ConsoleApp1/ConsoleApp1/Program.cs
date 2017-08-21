// 上海西口印刷有限公司 纸说信息部门 余峻峣 2017.8.20
// 固定资产明细格式变更
// 编辑环境VS2017 .Net框架 C#语言
using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples
{


    public class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("欢迎使用 每项输入完毕后回车提交\n\r" +
                              "请将目标文件放入程序同一目录中 \n\r" +
                              "运行此程序前请确保目标文件和输出文件处于关闭状态\n\r" +
                              "否则程序将出错且可能导致原数据丢失\n\r");
            Console.WriteLine("请输入 目标文件 文件名");
            Console.WriteLine("无需包含后缀名 目前仅支持xlsx");
            string filepath = Console.ReadLine();//输入文件目录
            string savepath = "输出.xlsx";//保存文件目录
            Console.WriteLine("请输入 需要转换的sheet的名称");
            string sheetname = Console.ReadLine();
            Console.WriteLine("请输入 表格模版的坐标\n\r左上到右下 中间以英文冒号隔开  例：G2:H6");
            string table = Console.ReadLine();
            //行列坐标表示
            int r, c;
            Console.WriteLine("请输入 表格模版左上角的行坐标  数字表示");
            r = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine("请输入 表格模版左上角的列坐标  数字表示");
            c = Convert.ToInt32(Console.ReadLine());
            using (XLWorkbook wb = new XLWorkbook(filepath+".xlsx"))
            {
                IXLWorksheet ws = wb.Worksheets.Worksheet(sheetname); //获得Sheet。

                //通过Worksheet(int position);函数获得sheet,position从1开始。

                //不能使用ws.RowCount();获得很大的值。
                //复制一份表格
                var rngTable = ws.Range(table);       //表格切片
                var excelTable = rngTable.CreateTable();//切片转化为表格

                //遍历所有可使用的行
                var rows = ws.RowsUsed();
                foreach (var row in rows)
                {
                    //检查是否是标题行
                    //是的话跳过此行
                    if (row.Cell(1).Value.ToString() == "公司名称")
                    {
                        //Console.WriteLine(row.Cell(1).Value.ToString());
                        continue;
                    }
                    ws.Cell(r, c).Value = excelTable; //插入复制的表格模板

                    //保存更改


                    //遍历所有的Cells

                    foreach (var cell in row.Cells())
                    {                   
                        string tempst = cell.Value.ToString();
                        if (tempst == "")
                        {
                            break;
                        }
                        Console.WriteLine(tempst);
                        ws.Cell(r, c+1).Value = tempst;
                        r += 1;//写下一列数据
                    }
                    r += 1;//每个表之间空一行
                }
                wb.SaveAs(savepath);//保存结果
            }
            Console.WriteLine("转换完毕 按任意键退出 输出文件初次打开可能需要excel的自动恢复");
            Console.ReadKey();
        }
    }
}