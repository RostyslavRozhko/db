﻿using Microsoft.Win32;
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
        private static String accessPath = "D:\\db.accdb";
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
                    foreach (String file in openFileDialog1.FileNames)
                    {
                        MessageBox.Show(file);
                        ExcelParser parser = new ExcelParser(file);

                        access.insertTeachers(parser.getTeachers());
                        access.insertWeeks(parser.getWeeks());
                        access.insertSchedule(parser.getYear(), parser.getSpeciality(), parser.getEntities());
                    }
                    MessageBox.Show("Імпортовано");
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
            Console.WriteLine("okay");
            Import();
        }

        private void mSubmit_Click(object sender, RoutedEventArgs e)
        {

        }

        private void tSubmit_Click(object sender, RoutedEventArgs e)
        {

        }

        private void sSubmit_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
