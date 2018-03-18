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
            DeclareStyles();
        }

        public void WriteHeader()
        {
            Worksheet.Cells[1, 2] = "Час";
            Worksheet.Cells[1, 3] = "Аудиторія";
            Worksheet.Cells[1, 4] = "Предмет";
            Worksheet.Cells[1, 5] = "Тип";
            Worksheet.Cells[1, 6] = "Спеціальність";
            Worksheet.Cells[1, 7] = "Рік гавчання";
            Worksheet.Cells[1, 8] = "Група";
            Worksheet.Cells[1, 9] = "Тиждень";
        }

        public void WriteData(List<String[]> data)
        {
            WriteHeader();

            int row = 2;
            String prevDay = "";
            int prevDayPos = -1;

            String prevTime = "";
            int prevTimePos = -1;

            SetStyle("Document", "A" + 1, "I" + data.Count + 1);

            foreach (String[] values in data)
            {
                if (values[0] == prevDay)
                {
                    Worksheet.Cells[row, 1] = "";

                    SetStyle("DayOfWeek", "A" + prevDayPos, "A" + row);
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
        }

        private void DeclareStyles()
        {
            Excel.Style Document = Workbook.Styles.Add("Document");
            Document.Font.Name = "Times New Roman";
            Document.Font.Size = 12;
            Document.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Document.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;


            Excel.Style DayOfWeek = Workbook.Styles.Add("DayOfWeek");
            Document.Font.Name = "Times New Roman";
            Document.Font.Size = 12;
            DayOfWeek.Orientation = Excel.XlOrientation.xlUpward;
            Document.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            Document.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }

        private void SetStyle(String StyleName, String Start, String End)
        {
            Excel.Range rangeStyles = Worksheet.get_Range(Start, End);
            rangeStyles.Style = StyleName;
        }

        public void Save()
        {
            Workbook.SaveAs("file.xlsx");
        }

        public void Close()
        {
            Workbook.Close(true);
            App.Quit();
        }
    }
}
