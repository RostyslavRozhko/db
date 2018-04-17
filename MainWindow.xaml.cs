using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Serialization;

namespace DBProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static Access access;

        public MainWindow()
        {
            InitializeComponent();
            access = new Access();
            String[] specs = access.GetSpecs();
            int c = 0;
            foreach (String spec in specs)
            {
                sSpecial.Items.Insert(c, spec);
            }

            errors.Text = access.getErrors("Цілісність1") + access.getErrors("Цілісність2");
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            access.deleteTables();
        }

        private bool Import()
        {
            access.deleteTables();
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog1.Multiselect = true;
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == true)
            {
                try
                {
                    MessageBox.Show("Імпорт може тривати кілька хвилин. Ми повідомимо, коли все буде готово.", "Обробка");

                    ExcelParser parser = new ExcelParser(openFileDialog1.FileNames);
                    parser.parseSchedule();

                    access.insertTeachers(parser.getTeachers());
                    access.insertSchedule(parser.getEntities());
                    access.insertWeeks(parser.getWeeks());

                    MessageBox.Show("Успішно імпортовано. Перевірте вкладку 'Помилки', щоб переконатися, що в розкладі немає накладок.", "Готово");
                    errors.Text = access.getErrors("Цілісність1") + access.getErrors("Цілісність2");
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            Import();
        }

        private void mSubmit_Click(object sender, RoutedEventArgs e)
        {
            bool mNCR = mLecRooms.IsChecked ?? false;
            bool mCR = mCompRooms.IsChecked ?? false;
            string mBld = mBuilding.Text;
            string mR = mRoom.Text;
            bool mAW = mAllWeeks.IsChecked ?? false;
            string mW = mWeek.Text;
            try
            {
                if (!mAW)
                {
                    ExcelWriter writer = new ExcelWriter();
                    String weekNum = mW;

                    String conditions = "WHERE ";
                    if (mR != "")
                    {
                        conditions += "Суб.Номер_аудиторії = '" + mR + "' ";
                    }
                    else if (mBld != "")
                    {
                        conditions += "Аудиторії.Корпус = '" + mBld + "' ";
                        if (mCR && !mNCR)
                        {
                            conditions += "AND Аудиторії.Компютерний_клас = 1 ";
                        }
                        else if (mNCR && !mCR)
                        {
                            conditions += "AND Аудиторії.Компютерний_клас = 0 ";
                        }
                    }
                    else if (mCR && !mNCR)
                    {
                        conditions += "Аудиторії.Компютерний_клас = 1 ";
                    }
                    else if (mNCR && !mCR)
                    {
                        conditions += "Аудиторії.Компютерний_клас = 0 ";
                    }

                    if (conditions == "WHERE ")
                    {
                        conditions = "";
                    }

                    Console.WriteLine(conditions);

                    List<String[]> data = access.Select("Методист2", conditions, 5, weekNum);
                    writer.WriteDataMeth(data, 2);
                    writer.Save();
                    writer.Close();
                }
                else
                {
                    ExcelWriter writer = new ExcelWriter();

                    String conditions = "WHERE ";
                    if (mR != "")
                    {
                        conditions += "Аудиторії.Номер_аудиторії = '" + mR + "' ";
                    }
                    else if (mBld != "")
                    {
                        conditions += "Аудиторії.Корпус = '" + mBld + "' ";
                        if (mCR && !mNCR)
                        {
                            conditions += "AND Аудиторії.Компютерний_клас = 1 ";
                        }
                        else if (mNCR && !mCR)
                        {
                            conditions += "AND Аудиторії.Компютерний_клас = 0 ";
                        }
                    }
                    else if (mCR && !mNCR)
                    {
                        conditions += "Аудиторії.Компютерний_клас = 1 ";
                    }
                    else if (mNCR && !mCR)
                    {
                        conditions += "Аудиторії.Компютерний_клас = 0 ";
                    }

                    if (conditions == "WHERE ")
                    {
                        conditions = "";
                    }

                    Console.WriteLine(conditions);
                    List<String[]> data = access.Select("Методист1", conditions, 5);

                    writer.WriteDataMeth(data, 2);
                    writer.Save();
                    writer.Close();
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void tSubmit_Click(object sender, RoutedEventArgs e)
        {
            string tLN = tLastname.Text.Replace("'", "’");
            bool tAW = tAllWeeks.IsChecked ?? false;
            string tW = tWeek.Text;
            try
            {
                if (!tAW)
                {
                    ExcelWriter writer = new ExcelWriter();
                    String conditions = "WHERE Викладачі.Прізвище LIKE '%" + tLN + "%' AND Розклад_Тижні.Номер_Тижні = " + tW;
                    List<String[]> data = access.Select("Викладач2", conditions, 8);

                    String[] header = { "", "Час", "Аудиторія", "Предмет", "Тип", "Спеціальність", "Рік навчання", "Група" };
                    writer.WriteHeader(header);
                    writer.WriteDataTeacher(data);
                    writer.Save();
                    writer.Close();
                }
                else
                {
                    ExcelWriter writer = new ExcelWriter();
                    String conditions = "WHERE Викладачі.Прізвище LIKE '%" + tLN + "%'";
                    List<String[]> data = access.Select("Викладач1", conditions, 9);

                    String[] header = { "", "Час", "Аудиторія", "Предмет", "Тип", "Спеціальність", "Рік навчання", "Група", "Тижні" };
                    writer.WriteHeader(header);
                    writer.WriteDataTeacher(data);
                    writer.Save();
                    writer.Close();
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void sSubmit_Click(object sender, RoutedEventArgs e)
        {
            string sS = sSpecial.SelectedItem.ToString();
            string sY = sYear.Text;
            bool sAW = sAllWeeks.IsChecked ?? false;
            string sW = sWeek.Text;
            Console.WriteLine(sS);

            try
            {
                if (!sAW)
                {
                    ExcelWriter writer = new ExcelWriter();
                    String conditions = "WHERE Розклад.Спеціальність = '" + sS + "' AND Розклад.Рік_навчання = " + sY + " AND Розклад_Тижні.Номер_Тижні = " + sW;
                    List<String[]> data = access.Select("Студент2", conditions, 7);

                    String[] header = { "", "Час", "Аудиторія", "Викладач", "Предмет", "Тип", "Група" };
                    writer.WriteHeader(header);
                    writer.WriteDataTeacher(data);
                    writer.Save();
                    writer.Close();
                }
                else
                {
                    ExcelWriter writer = new ExcelWriter();
                    String conditions = "WHERE Розклад.Спеціальність = '" + sS + "' AND Розклад.Рік_навчання = " + sY;
                    List<String[]> data = access.Select("Студент1", conditions, 8);

                    String[] header = { "", "Час", "Аудиторія", "Викладач", "Предмет", "Тип", "Група", "Тижні" };
                    writer.WriteHeader(header);
                    writer.WriteDataTeacher(data);
                    writer.Save();
                    writer.Close();
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void mRoom_TextChanged(object sender, TextChangedEventArgs e)
        {
            mLecRooms.IsChecked = false;
            mCompRooms.IsChecked = false;
            mBuilding.Text = "";
        }

        private void mWeek_TextChanged(object sender, TextChangedEventArgs e)
        {
            mAllWeeks.IsChecked = false;
        }

        private void tWeek_TextChanged(object sender, TextChangedEventArgs e)
        {
            tAllWeeks.IsChecked = false;
        }

        private void sWeek_TextChanged(object sender, TextChangedEventArgs e)
        {
            sAllWeeks.IsChecked = false;
        }
    }
}
