using System;
using System.Collections.Generic;
using System.Data.OleDb;
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

namespace DBProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;

            InitializeComponent();
            String path = "C:\\Work\\c#\\db\\bin\\Debug\\Інформатика_3_весна_17_18н_р_Microsoft.xlsx";
            ExcelParser parser = new ExcelParser(path);

            String accessPath = @"C:\Work\c#\db\bin\Debug\db.accdb";
            Access access = new Access(accessPath);

            access.deleteTables();
            access.insertTeachers(parser.getTeachers());
            access.insertWeeks(parser.getWeeks());
            access.insertSchedule(parser.getYear(), parser.getSpeciality(), parser.getEntities());
        }

        private void MenuItem_MouseDown(object sender, MouseButtonEventArgs e)
        {

        }

        private void ClearButton_MouseDown(object sender, MouseButtonEventArgs e)
        {
            grid.Children.Clear();
            String path = @"D:\db\2017_2018__Spring\Інформатика  -3 весна 17-18н.р.  Microsoft Office Excel.xlsx";
            ExcelParser parser = new ExcelParser(path);

            String accessPath = "D:\\GIT\\dbproject\\db.accdb";
            Access access = new Access(accessPath);
            Console.WriteLine("suka");
            access.deleteTables();
        }
    }
}
