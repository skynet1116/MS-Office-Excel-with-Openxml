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
            string filepath = @"C:\Users\15752\Desktop\固定资产编号明细.xlsx";//输入文件目录
            string savepath = @"C:\Users\15752\Desktop\输出.xlsx";//保存文件目录
            //string filepath = @"C:\Users\Administrator\Desktop\固定资产编号明细(1).xlsx";
            using (XLWorkbook wb = new XLWorkbook(filepath))
            {
                IXLWorksheet ws = wb.Worksheets.Worksheet("资产明细"); //获得第一个Sheet。

                //通过Worksheet(int position);函数获得sheet,position从1开始。

                //不能使用ws.RowCount();获得很大的值。
                //复制一份表格
                var rngTable = ws.Range("G2:H6");       //表格切片
                var excelTable = rngTable.CreateTable();//切片转化为表格
                //行列坐标表示
                int r = 2;
                int c = 7;
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
                        ws.Cell(r, 8).Value = tempst;
                        r += 1;//写下一列数据
                    }
                    r += 1;//每个表之间空一行
                }
                wb.SaveAs(savepath);//保存结果
            }
        }
    }
}