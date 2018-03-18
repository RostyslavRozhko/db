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

        public List<String[]> selectTeacher()
        {
            List<String[]> result = new List<String[]>();
            OleDbCommand cmd;
            OleDbDataReader RS;
            using (OleDbConnection Connection = new OleDbConnection())
            {

                Connection.ConnectionString = connectionstring;
                Connection.Open();
                String sql = "SELECT Дні_тижня.День_назва AS День, Пара_час_з & ' - ' & Пара_час_до AS Пара, Розклад.Аудиторія, Розклад.Предмет, Розклад.Тип, Розклад.Спеціальність, Розклад.Рік_навчання, Розклад.Група FROM(((Розклад INNER JOIN Дні_тижня ON Розклад.День = Дні_тижня.День_номер) INNER JOIN Пари ON Розклад.Пара_номер = Пари.Пара_номер) INNER JOIN Викладачі ON Розклад.Викладач = Викладачі.Викладач_код) WHERE Викладачі.Прізвище LIKE '%Сініцина%' ORDER BY День, Розклад.Пара_номер, Аудиторія, Розклад.Група ";
                cmd = new OleDbCommand(sql, Connection);
                RS = cmd.ExecuteReader();
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
                RS.Close();
            }
            return result;
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

        public void insertSchedule(List<ExcelRecord> records)
        {
            foreach (ExcelRecord record in records)
            {
                OleDbCommand cmd;
                OleDbDataReader RS;

                String sql = "INSERT INTO Розклад (Спеціальність, Рік_навчання, Номер_запису, День, Пара_номер, Аудиторія, Предмет, Група, Викладач) " +
                    "VALUES (" + record.ToString() + ")";
                Console.WriteLine(sql);
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
