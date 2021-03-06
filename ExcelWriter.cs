﻿using Microsoft.Win32;
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
        private int row = 1;

        public ExcelWriter()
        {
            App = new Excel.Application();
            Workbook = App.Workbooks.Add();
            Worksheet = (Excel.Worksheet) Workbook.Worksheets.get_Item(1);
        }

        public void WriteHeader(String[] arr, int concat = 1)
        {
            for(int i = 0; i < arr.Length; i++)
            {
                Worksheet.Cells[1, i+1] = arr[i];

            }
            Worksheet.Range[Worksheet.Cells[row, 1], Worksheet.Cells[row, concat]].Merge();
            row++;
        }

        public void WriteData(List<String[]> data)
        {
            if (data.Count == 0)
                return;

            int rows = data.Count + 1;
            int cols = data[0].Length;

            Excel.Range table = Worksheet.Range[Worksheet.Cells[1, 4], Worksheet.Cells[rows, cols]];
            table.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            table.NumberFormat = "@";

            String prevDay = "";
            int prevDayPos = -1;

            String prevTime = "";
            int prevTimePos = -1;

            String prevRoom = "";
            int prevRoomPos = -1;

            foreach (String[] values in data)
            {
                if (values[0] == prevDay)
                {
                    Worksheet.Cells[row, 1] = "";
                    Worksheet.Range[Worksheet.Cells[prevTimePos, 1], Worksheet.Cells[row, 1]].Merge();
                }
                else
                {
                    prevTimePos = row;
                    prevTime = "";
                    Worksheet.Cells[row, 1] = values[0];
                }

                if (values[1] == prevTime && values[0] == prevDay)
                {
                    Worksheet.Cells[row, 2] = "";
                    Worksheet.Range[Worksheet.Cells[prevTimePos, 2], Worksheet.Cells[row, 2]].Merge();
                }
                else
                {
                    prevTimePos = row;
                    prevRoom = "";
                    Worksheet.Cells[row, 2] = values[1];
                }

                if (values[2] == prevRoom)
                {
                    Worksheet.Cells[row, 3] = "";
                    Worksheet.Range[Worksheet.Cells[prevRoomPos, 3], Worksheet.Cells[row, 3]].Merge();
                }
                else
                {
                    prevRoomPos = row;
                    Worksheet.Cells[row, 3] = values[2];
                }

                for (int i = 3; i < cols; i++)
                {
                    Worksheet.Cells[row, i + 1] = values[i];
                }
                prevDay = values[0];
                prevTime = values[1];
                prevRoom = values[2];
                row++;
            }
            SetStyles(rows, cols);

        }

        public void WriteDataTeacher(List<String[]> data)
        {
            if (data.Count == 0)
                return;

            int rows = data.Count + 1;
            int cols = data[0].Length;

            Excel.Range table = Worksheet.Range[Worksheet.Cells[1, 3], Worksheet.Cells[rows, cols]];
            table.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            table.NumberFormat = "@";

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

                if (values[1] == prevTime  && values[0] == prevDay)
                {
                    Worksheet.Cells[row, 2] = "";
                    Worksheet.Range[Worksheet.Cells[prevTimePos, 2], Worksheet.Cells[row, 2]].Merge();
                }
                else
                {
                    prevTimePos = row;
                    Worksheet.Cells[row, 2] = values[1];
                }

                for (int i = 2; i < cols; i++)
                {
                    Worksheet.Cells[row, i + 1] = values[i];
                }
                prevDay = values[0];
                prevTime = values[1];
                row++;
            }
            SetStyles(rows, cols);
        }

        public void WriteDataMeth(List<String[]> data, int subCols = 2)
        {
            if (data.Count == 0)
                return;

            int rows = data.Count + 1;
            int cols = data[0].Length;

            Excel.Range table = Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[rows / 5, 2 + subCols * 6]];
            table.NumberFormat = "@";

            String prevDay = "";
            String prevTime = "";
            String prevRoom = "";

            int dayCol = 3 - subCols;

            int row = 2;

            foreach (String[] values in data)
            {
                if (values[0] != prevDay)
                {
                    dayCol += subCols;
                    Worksheet.Cells[1, dayCol] = values[0];
                    row = 2;
                }

                Worksheet.Cells[row, 1] = values[1];
                Worksheet.Cells[row, 2] = values[2];

                if (values[0] == prevDay
                    && values[1] == prevTime
                    && values[2] == prevRoom)
                {
                    for (int i = 3, counter = 0; i < cols; i++)
                    {
                        Worksheet.Cells[row - 1, dayCol + counter] =
                            (string)(Worksheet.Cells[row - 1 , dayCol + counter] as Excel.Range).Value  
                            + " / " 
                            + values[i];
                        counter++;
                    }
                } else
                {
                    for (int i = 3, counter = 0; i < cols; i++)
                    {
                        Worksheet.Cells[row, dayCol + counter] = values[i];
                        counter++;
                    }

                    prevDay = values[0];
                    prevTime = values[1];
                    prevRoom = values[2];
                    row++;
                }
            }
            setStylesMeth(row - 1, 2 + subCols * 6);
        }

        private void setStylesMeth(int rows, int cols)
        {
            Worksheet.Columns.Font.Name = "Times New Roman";
            Worksheet.Columns.Font.Size = 9;
            Worksheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


            Excel.Range table = Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[rows, cols]];
            table.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            for (int i = 1; i <= cols; i++)
            {
                Excel.Range col = Worksheet.Range[Worksheet.Cells[1, i], Worksheet.Cells[rows, i]];
                col.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                col.WrapText = true;
                col.ColumnWidth = 9;
                col.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
                col.Columns.AutoFit();
            }
        }

        private void SetStyles(int rows, int cols)
        {
            Excel.Range table = Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[rows, cols]];
            table.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Worksheet.Columns.Font.Name = "Times New Roman";
            Worksheet.Columns.Font.Size = 12;
            Worksheet.Columns.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

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

            Excel.Range group = Worksheet.Range[Worksheet.Cells[1, 7], Worksheet.Cells[rows, 7]];
            group.Columns.AutoFit();

            Excel.Range week = Worksheet.Range[Worksheet.Cells[1, 8], Worksheet.Cells[rows, 8]];
            week.Columns.AutoFit();

            Excel.Range col = Worksheet.Range[Worksheet.Cells[1, 9], Worksheet.Cells[rows, 9]];
            col.Columns.AutoFit();
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
