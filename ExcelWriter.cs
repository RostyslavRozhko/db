using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace DBProject
{
    class ExcelWriter
    {
        private static Excel.Workbook Workbook = null;
        private static Excel.Application App = null;
        private static Excel.Worksheet Worksheet = null;

        public ExcelWriter()
        {
            App = new Excel.Application();
            Workbook = App.Workbooks.Add();
            Worksheet = (Excel.Worksheet) Workbook.Worksheets.get_Item(1);
        }

        public void WriteHeader(String[] arr)
        {
            for(int i = 0; i < arr.Length; i++)
            {
                Worksheet.Cells[1, i+1] = arr[i];
            }
        }

        public void WriteData(List<String[]> data)
        {
            if (data.Count == 0)
                return;

            int row = 2;
            String prevDay = "";
            int prevDayPos = -1;

            String prevTime = "";
            int prevTimePos = -1;

            foreach (String[] values in data)
            {
                if (values[0] == prevDay)
                {
                    Worksheet.Cells[row, 1] = "";
                    Worksheet.Range[Worksheet.Cells[prevDayPos, 1], Worksheet.Cells[row, 1]].Merge();
                }
                else
                {
                    prevDayPos = row;
                    Worksheet.Cells[row, 1] = values[0];
                }

                if (values[1] == prevTime)
                {
                    Worksheet.Cells[row, 2] = "";
                    Worksheet.Range[Worksheet.Cells[prevTimePos, 2], Worksheet.Cells[row, 2]].Merge();
                }
                else
                {
                    prevTimePos = row;
                    Worksheet.Cells[row, 2] = values[1];
                }

                for (int i = 2; i < 8; i++)
                {
                    Worksheet.Cells[row, i + 1] = values[i];
                }
                prevDay = values[0];
                prevTime = values[1];
                row++;
            }
            SetStyles(data.Count+1, data[0].Length+1);
        }

        private void SetStyles(int rows, int cols)
        {
            Worksheet.Columns.Font.Name = "Times New Roman";
            Worksheet.Columns.Font.Size = 12;
            Worksheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            Excel.Range table = Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[rows, cols]];
            table.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.Range head = Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, cols]];
            head.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            head.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
            head.Columns.AutoFit();

            Excel.Range day = Worksheet.Range[Worksheet.Cells[2, 1], Worksheet.Cells[rows, 1]];
            day.Orientation = Excel.XlOrientation.xlUpward;
            day.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            day.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            day.Columns.AutoFit();

            Excel.Range time = Worksheet.Range[Worksheet.Cells[2, 2], Worksheet.Cells[rows, 2]];
            time.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            time.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            time.Columns.AutoFit();

            Excel.Range title = Worksheet.Range[Worksheet.Cells[2, 4], Worksheet.Cells[rows, 4]];
            title.Font.Bold = true;
            title.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            title.Columns.AutoFit();

            Excel.Range type = Worksheet.Range[Worksheet.Cells[2, 5], Worksheet.Cells[rows, 5]];
            type.Columns.AutoFit();

            Excel.Range spec = Worksheet.Range[Worksheet.Cells[2, 6], Worksheet.Cells[rows, 6]];
            spec.Columns.AutoFit();
        }

        public void Save()
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();

            saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveFileDialog1.FilterIndex = 2;
            saveFileDialog1.RestoreDirectory = true;

            if (saveFileDialog1.ShowDialog() == true)
            {
                Workbook.SaveAs(saveFileDialog1.FileName);
            }
        }

        public void Close()
        {
            Workbook.Close(true);
            App.Quit();
        }
    }
}
