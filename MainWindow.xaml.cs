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
        private static String accessPath = @"D:\db.accdb";
        private static Access access = new Access(accessPath);

        public MainWindow()
        {
            InitializeComponent();
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

                    access.insertTeachers(parser.getTeachers());
                    access.insertSchedule(parser.getEntities());
                    access.insertWeeks(parser.getWeeks());

                    MessageBox.Show("Успішно імпортовано. Перевірте вкладку 'Помилки', щоб переконатися, що в розкладі немає накладок.", "Готово");
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
            bool mAR = mAllRooms.IsChecked ?? false;
            bool mCR = mCompRooms.IsChecked ?? false;
            string mBld = mBuilding.Text;
            string mR = mRoom.Text;
            bool mAW = mAllWeeks.IsChecked ?? false;
            string mW = mWeek.Text;
            Console.WriteLine(mAR + " " + mCR + " ");
            try
            {
                ExcelWriter writer = new ExcelWriter();

                String[] header = { };
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void tSubmit_Click(object sender, RoutedEventArgs e)
        {
            string tLN = tLastname.Text;
            bool tAW = tAllWeeks.IsChecked ?? false;
            string tW = tWeek.Text;
            Console.WriteLine(tLN + " " + tW + " ");
            try
            {
                if(!tAW)
                {
                    ExcelWriter writer = new ExcelWriter();
                    String conditions = "WHERE Викладачі.Прізвище LIKE '%" + tLN + "%' AND Розклад_Тижні.Номер_Тижні = " + tW;
                    List<String[]> data = access.SelectTeacher("Викладач2", conditions, 8);

                    String[] header = { "", "Час", "Аудиторія", "Предмет", "Тип", "Спеціальність", "Рік навчання", "Група" };
                    writer.WriteHeader(header);
                    writer.WriteDataTeacher(data);
                    writer.Save();
                    writer.Close();
                } else
                {
                    ExcelWriter writer = new ExcelWriter();
                    String conditions = "WHERE Викладачі.Прізвище LIKE '%" + tLN + "%'";
                    List<String[]> data = access.SelectTeacher("Викладач1", conditions, 9);

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
            string sS = sSpecial.Text;
            string sY = sYear.Text;
            bool sAW = sAllWeeks.IsChecked ?? false;
            string sW = sWeek.Text;
        }

        private void mBuilding_TextChanged(object sender, TextChangedEventArgs e)
        {
            mAllRooms.IsChecked = false;
        }

        private void mRoom_TextChanged(object sender, TextChangedEventArgs e)
        {
            mAllRooms.IsChecked = false;
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
