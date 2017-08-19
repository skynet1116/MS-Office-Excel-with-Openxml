
using ClosedXML.Excel;
using System;

namespace ClosedXML_Examples
{


    public class Program
    {
        static void Main(string[] args)
        {
            using (XLWorkbook wb = new XLWorkbook(@"C:\Users\Administrator\Desktop\固定资产编号明细(1).xlsx"))
            {
                IXLWorksheet ws = wb.Worksheets.Worksheet("资产明细"); //获得第一个Sheet。

                //通过Worksheet(int position);函数获得sheet,position从1开始。

                //不能使用ws.RowCount();获得很大的值。
                //复制一份表格
                var rngTable = ws.Range("G2:H7");       //表格切片
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
                        Console.WriteLine(row.Cell(1).Value.ToString());
                        continue;
                    }
                    r += 6;  //表格之间相隔6行
                    ws.Cell(r, c).Value = excelTable;

                    //保存更改


                    //遍历所有的Cells

                    foreach (var cell in row.Cells())

                    {

                        string tempst = cell.Value.ToString();
                        Console.WriteLine(tempst);

                    }

                }
                wb.SaveAs(@"C:\Users\Administrator\Desktop\固定资产编号明细(1).xlsx");
                /*string filepath = @"C:\Users\Administrator\Desktop\固定资产输出.xlsx";
                var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("Sample Sheet");
                // From a list of strings
                var listOfStrings = new List<String>();
                listOfStrings.Add("House");
                listOfStrings.Add("Car");
                ws.Cell(1, 1).Value = "Strings";
                //ws.Cell(1, 1).AsRange().AddToNamed("Titles");
                ws.Cell(2, 1).Value = listOfStrings;
                workbook.SaveAs(filepath);*/
            }
        }


    }
}