using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace DBProject
{
    class Access
    {
        private OleDbConnection Connection;
        private Dictionary<String, String> Queries;

        public Access(String path)
        {
            String connectionstring = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Persist Security Info=False";
            Connection = new OleDbConnection();
            Connection.ConnectionString = connectionstring;
            Connection.Open();

            GetQueries();
        }

        public void Close()
        {
            Connection.Close();
        }

        private void GetQueries()
        {
            Queries = new Dictionary<String, String>();

            OleDbCommand cmd = new OleDbCommand("SELECT Назва, Запит FROM Запити", Connection);
            OleDbDataReader RS = cmd.ExecuteReader();
            while (RS.Read())
            {
                Queries.Add(RS[0].ToString(), RS[1].ToString());
            }
            RS.Close();
        }

        public List<String[]> SelectTeacher(String queryName, String conditions)
        {
            List<String[]> result = new List<String[]>();

            String query = Queries["queryName"];
            String sql = query.Replace("$", conditions);

            OleDbCommand cmd = new OleDbCommand(sql, Connection);
            OleDbDataReader RS = cmd.ExecuteReader();

            while (RS.Read())
            {
                String[] array = new String[8];
                array[0] = RS[0].ToString();
                array[1] = RS[1].ToString();
                array[2] = RS[2].ToString();
                array[3] = RS[3].ToString();
                array[4] = RS[4].ToString();
                array[5] = RS[5].ToString();
                array[6] = RS[6].ToString();
                array[7] = RS[7].ToString();
                result.Add(array);
            }

            return result;
        }

        public void insertTeachers(List<Teacher> teachers)
        {
            foreach (Teacher teacher in teachers)
            {
                String sql = "INSERT INTO Викладачі (Викладач_код, Прізвище, Ініціали, Посада) VALUES " + teacher.ToString();

                OleDbCommand cmd = new OleDbCommand(sql, Connection);
                OleDbDataReader RS = cmd.ExecuteReader();
                RS.Close();
            }
        }

        public void insertWeeks(List<Weeks> weeks)
        {
            foreach (Weeks weeksList in weeks)
            {
                foreach (String weekNum in weeksList.WeeksList)
                {
                    String sql = "INSERT INTO Розклад_Тижні (Номер_запису_Розклад, Номер_Тижні) VALUES (" + weeksList.EntityId + ", " + weekNum + ")";

                    OleDbCommand cmd = new OleDbCommand(sql, Connection);
                    OleDbDataReader RS = cmd.ExecuteReader();
                    RS.Close();
                }
            }
        }

        public void insertSchedule(List<ExcelRecord> records)
        {
            foreach (ExcelRecord record in records)
            {
                String sql = "INSERT INTO Розклад (Спеціальність, Рік_навчання, Номер_запису, День, Пара_номер, Аудиторія, Предмет, Група, Викладач) " +
                    "VALUES (" + record.ToString() + ")";

                OleDbCommand cmd = new OleDbCommand(sql, Connection);
                OleDbDataReader RS = cmd.ExecuteReader();
                RS.Close();
            }
        }

        public void deleteTables()
        {
            try
            {
                String[] commands = {
                    "DELETE FROM Викладачі",
                    "DELETE FROM Розклад",
                    "DELETE FROM Розклад_Тижні"
            };

                foreach (String sql in commands)
                {
                    OleDbCommand cmd = new OleDbCommand(sql, Connection);
                    OleDbDataReader RS = cmd.ExecuteReader();
                    RS.Close();

                }
            } catch(Exception e)
            {
                Console.WriteLine(e);
            }
            
        }
    }
}
