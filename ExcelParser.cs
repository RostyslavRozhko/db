using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace DBProject
{
    class ExcelParser
    {
        private static Excel.Workbook Workbook = null;
        private static Excel.Application App = null;
        private static Excel.Worksheet Worksheet = null;
        private static int numberOfRows;
        private List<Teacher> Teachers;
        private List<ExcelRecord> Records;
        private List<Weeks> WeeksList;

        public ExcelParser(string[] pathToFiles)
        {
            Teachers = new List<Teacher>();
            Records = new List<ExcelRecord>();
            WeeksList = new List<Weeks>();

            foreach (String file in pathToFiles)
            {
                App = new Excel.Application();
                App.Visible = false;
                Workbook = App.Workbooks.Open(file);
                Workbook.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
                Worksheet = (Excel.Worksheet)Workbook.Sheets[1];
                numberOfRows = Worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                Parse();
                Close();
            }
        }

        public static void Close()
        {
            Workbook.Saved = true;
            App.Quit();

        }

        private void Parse()
        {
            String Year = getYear();
            String Speciality = getSpeciality();

            int DayOfWeek = 0;
            int Time = 0;

            for (int index = 11; index <= numberOfRows; index++)
            {
                Array Values = (Array)Worksheet.get_Range("A" + index.ToString(), "G" + index.ToString()).Cells.Value;

                if (Values.GetValue(1, 1) != null)
                {
                    DayOfWeek++;
                    Time = 0;
                }

                if (Values.GetValue(1, 2) != null)
                {
                    Time++;
                }

                if (Values.GetValue(1, 3) != null)
                {
                    String teacherId;
                    String groupId;
                    String teacherName = System.Security.SecurityElement.Escape(Values.GetValue(1, 4).ToString());

                    Teacher teacher = Teachers.Find(x => x.Name == teacherName);
                    if (teacher == null)
                    {
                        Teacher newTeacher = new Teacher(teacherName);
                        Teachers.Add(newTeacher);
                        teacherId = newTeacher.Id.ToString();
                    }
                    else
                    {
                        teacherId = teacher.Id.ToString();
                    }

                    if (Values.GetValue(1, 5).ToString() == "лекція")
                    {
                        groupId = "0";
                    }
                    else
                    {
                        groupId = Values.GetValue(1, 5).ToString();
                    }

                    String room = "NULL";
                    if (Values.GetValue(1, 7) != null)
                    {
                        room = Values.GetValue(1, 7).ToString().Replace(" ", "");
                    }
                    ExcelRecord entity = new ExcelRecord(
                            Year,
                            Speciality,
                            DayOfWeek.ToString(),
                            Time.ToString(),
                            System.Security.SecurityElement.Escape(Values.GetValue(1, 3).ToString()),
                            teacherId,
                            groupId,
                            room
                        );
                    Weeks weeksObj = new Weeks(entity.Id, weeks(Values.GetValue(1, 6).ToString()));
                    Records.Add(entity);
                    WeeksList.Add(weeksObj);
                }
            }
        }

        private List<String> weeks(String weeks)
        {

            string[] splitWeeks = weeks.Replace(" ", "").Split(',');
            List<String> result = new List<String>();
            for (int i = 0; i < splitWeeks.Length; i++)
            {
                string currElement = splitWeeks[i];
                if (currElement.IndexOf('-') > -1)
                {
                    string[] split = currElement.Split('-');
                    for (int j = Int32.Parse(split[0]); j < Int32.Parse(split[1]); j++)
                    {
                        result.Add(j.ToString());
                    }
                }
                else
                {
                    if(currElement != "")
                        result.Add(currElement);
                }
            }
            return result;
        }

        public String getFaculty()
        {
            return (String)(Worksheet.Cells[6, 1] as Excel.Range).Value;
        }

        public String getSpeciality()
        {
            String output = (String)(Worksheet.Cells[7, 1] as Excel.Range).Value;
            Regex regex = new Regex("\"([^\"]*)\"");
            Match match = regex.Match(output);
            return match.Value.Replace("\"", "");
        }

        public String getYear()
        {
            String output = (String)(Worksheet.Cells[7, 1] as Excel.Range).Value;
            Regex regex = new Regex("[1-4]");
            Match match = regex.Match(output);
            return match.Value;
        }

        public List<ExcelRecord> getEntities()
        {
            return Records;
        }

        public List<Teacher> getTeachers()
        {
            return Teachers;
        }

        public List<Weeks> getWeeks()
        {
            return WeeksList;
        }

    }
}
