using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;

namespace DBProject
{
    class Access
    {
        string connectionstring;
        public Access(String path)
        {
            this.connectionstring = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Persist Security Info=False";
        }

        public void select(String table)
        {
            OleDbCommand cmd;
            OleDbDataReader RS;
            using (OleDbConnection Connection = new OleDbConnection())
            {

                Connection.ConnectionString = connectionstring;
                Connection.Open();
                cmd = new OleDbCommand("SELECT * FROM " + table, Connection);
                RS = cmd.ExecuteReader();
                while (RS.Read())
                {
                    Console.WriteLine(RS[0] + " " + RS[1]);
                }
                RS.Close();
            }
        }

        public void insertTeachers(List<Teacher> teachers)
        {
            foreach (Teacher teacher in teachers)
            {
                OleDbCommand cmd;
                OleDbDataReader RS;

                String sql = "INSERT INTO Викладачі (Викладач_код, Прізвище, Ініціали, Посада) VALUES " + teacher.ToString();

                using (OleDbConnection Connection = new OleDbConnection())
                {

                    Connection.ConnectionString = connectionstring;
                    Connection.Open();
                    cmd = new OleDbCommand(sql, Connection);
                    RS = cmd.ExecuteReader();
                    RS.Close();
                }
            }
        }

        public void insertWeeks(List<Weeks> weeks)
        {
            foreach (Weeks weeksList in weeks)
            {
                foreach (String weekNum in weeksList.WeeksList)
                {
                    OleDbCommand cmd;
                    OleDbDataReader RS;

                    String sql = "INSERT INTO Розклад_Тижні (Номер_запису_Розклад, Номер_Тижні) VALUES (" + weeksList.EntityId + ", " + weekNum + ")";

                    using (OleDbConnection Connection = new OleDbConnection())
                    {

                        Connection.ConnectionString = connectionstring;
                        Connection.Open();
                        cmd = new OleDbCommand(sql, Connection);
                        RS = cmd.ExecuteReader();
                        RS.Close();
                    }
                }
            }
        }

        public void insertSchedule(String year, String speciality, List<ExcelRecord> records)
        {
            foreach (ExcelRecord record in records)
            {
                OleDbCommand cmd;
                OleDbDataReader RS;

                String sql = "INSERT INTO Розклад (Спеціальність, Рік_навчання, Номер_запису, День, Пара_номер, Аудиторія, Предмет, Група, Викладач) " +
                    "VALUES ('" + speciality + "', " + year + ", " + record.ToString() + ")";

                using (OleDbConnection Connection = new OleDbConnection())
                {

                    Connection.ConnectionString = connectionstring;
                    Connection.Open();
                    cmd = new OleDbCommand(sql, Connection);
                    RS = cmd.ExecuteReader();
                    RS.Close();
                }
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
                    OleDbCommand cmd;
                    OleDbDataReader RS;

                    using (OleDbConnection Connection = new OleDbConnection())
                    {

                        Connection.ConnectionString = connectionstring;
                        Connection.Open();
                        cmd = new OleDbCommand(sql, Connection);
                        RS = cmd.ExecuteReader();
                        RS.Close();
                    }
                }
            } catch(Exception e)
            {
                Console.WriteLine(e);
            }
            
        }
    }
}
