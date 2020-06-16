using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace WpfApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public string path;
        public string savePath;

        public MainWindow()
        {
            InitializeComponent();
        }

        List<GasComposition> myGas = new List<GasComposition>();

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {

        }

        private void MenuItem_Click_1(object sender, RoutedEventArgs e)
        {
            //calc
            if (gasNameTextBox.Text != "") //protection from empty calculation
            {
                GasComposition myTempGas = myGas[0]; //extract gas from collection
                GasComposition myGas1NormalizedComposition = GasComposition.Normalize(ref myTempGas); //normalize
                GasComposition myGas1Calculated = GasComposition.CalculateProperties(ref myGas1NormalizedComposition); //calc prop
                                                                                                                       // send RO and R to screen
                dencityTextBox.Text = myGas1Calculated.mixtureDencity.ToString();
                gasConstantTextBox.Text = myGas1Calculated.mixtureR.ToString();
            }
            else
            {
                MessageBox.Show("Нет данных для расчета", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MenuItem_Click_2(object sender, RoutedEventArgs e)
        {
            //clear
            CleanScreen();
        }

        private void ToggleButton_Click(object sender, RoutedEventArgs e)
        {
            //open
            //--clean previous data----
            CleanScreenNoMessage();
            //-----------------
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\temp";
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
                        if (openFileDialog1.ShowDialog() == true)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        path = openFileDialog1.FileName;
                        savePath = path;
                        GasComposition myGas1 = new GasComposition();
                        GasComposition myGas1Composition = GasComposition.ReadExcelFile(ref path, ref myGas1);

                        gasNameTextBox.Text = myGas1Composition.gasName;//gas name

                        for (int j = 0; j < myGas1Composition.size; j++)// по всем строкам
                        {
                            string cName = myGas1Composition.componentName[j];
                            string cFormula = myGas1Composition.componentFormula[j];
                            string cData0 = myGas1Composition.componentData[j, 0].ToString();
                            string cData1 = myGas1Composition.componentData[j, 1].ToString();
                            //string cWeight = myGas1Composition.componentWeight[j].ToString();
                            gasListView.Items.Add(new
                            {
                                Number = j + 1,
                                componentName = cName,
                                componentFormula = cFormula,
                                componentMolarWeight = cData0,
                                componentVolume = cData1
                            });


                        }

                        myGas.Add(myGas1); //send gas to collection
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk:" + ex.Message);
                }

            }

        }

        private void ToggleButton_Click_1(object sender, RoutedEventArgs e)
        {
            //save
            if (gasConstantTextBox.Text != "" && dencityTextBox.Text != "") //if any data to save
            {
                GasComposition dataForSaving = myGas[0]; //extract gas from collection
                GasComposition.SaveResultsToExcel(ref savePath, ref dataForSaving);
            }
            else
            {
                MessageBox.Show("Нет данных для сохранения", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ToggleButton_Click_2(object sender, RoutedEventArgs e)
        {
            //calc
            if (gasNameTextBox.Text != "") //protection from empty calculation
            {
                GasComposition myTempGas = myGas[0]; //extract gas from collection
                GasComposition myGas1NormalizedComposition = GasComposition.Normalize(ref myTempGas); //normalize
                GasComposition myGas1Calculated = GasComposition.CalculateProperties(ref myGas1NormalizedComposition); //calc prop
                                                                                                                       // send RO and R to screen
                dencityTextBox.Text = myGas1Calculated.mixtureDencity.ToString();
                gasConstantTextBox.Text = myGas1Calculated.mixtureR.ToString();
            }
            else
            {
                MessageBox.Show("Нет данных для расчета", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ToggleButton_Click_3(object sender, RoutedEventArgs e)
        {
            //clear
            CleanScreen();
        }

        private void MenuItem_Click_3(object sender, RoutedEventArgs e)
        {
            //open
            //--clean previous data----
            CleanScreenNoMessage();
            //-----------------
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"C:\temp";
            // openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            //if (openFileDialog1.ShowDialog() == DialogResult.OK)
            if (openFileDialog1.ShowDialog() == true)
                {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        path = openFileDialog1.FileName;
                        savePath = path;
                        GasComposition myGas1 = new GasComposition();
                        GasComposition myGas1Composition = GasComposition.ReadExcelFile(ref path, ref myGas1);

                        gasNameTextBox.Text = myGas1Composition.gasName;//gas name

                        for (int j = 0; j < myGas1Composition.size; j++)// по всем строкам
                        {
                            string cName = myGas1Composition.componentName[j];
                            string cFormula = myGas1Composition.componentFormula[j];
                            string cData0 = myGas1Composition.componentData[j, 0].ToString();
                            string cData1 = myGas1Composition.componentData[j, 1].ToString();
                            //string cWeight = myGas1Composition.componentWeight[j].ToString();
                            gasListView.Items.Add(new { Number =j+1 , componentName = cName, 
                            componentFormula = cFormula, componentMolarWeight = cData0, componentVolume = cData1 });
                           
                        }

                        myGas.Add(myGas1); //send gas to collection
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk:" + ex.Message);
                }

            }
        }

        private void MenuItem_Click_4(object sender, RoutedEventArgs e)
        {
            //save
            if (gasConstantTextBox.Text != "" && dencityTextBox.Text != "") //if any data to save
            {
                GasComposition dataForSaving = myGas[0]; //extract gas from collection
                GasComposition.SaveResultsToExcel(ref savePath, ref dataForSaving);
            }
            else
            {
                MessageBox.Show("Нет данных для сохранения", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void MenuItem_Click_5(object sender, RoutedEventArgs e)
        {
            //Exit
            this.Close();
        }

        private void MenuItem_Click_6(object sender, RoutedEventArgs e)
        {
            // Content
            DescriptionWindow DescrWindow = new DescriptionWindow();
            DescrWindow.Show();
        }

        private void MenuItem_Click_7(object sender, RoutedEventArgs e)
        {
            //Author
            AuthorWindow AuWindow = new AuthorWindow();
            AuWindow.Show();
        }


        public void CleanScreen()
        {
            if (gasConstantTextBox.Text != "" && dencityTextBox.Text != "") //protection from empty cleaning
            {
                gasNameTextBox.Text = gasConstantTextBox.Text = dencityTextBox.Text = "";
                gasListView.Items.Clear(); //clear screen 
                myGas.Clear(); //clear gas 
            }
            else
            {
                MessageBox.Show("Форма уже очищена", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        public void CleanScreenNoMessage()
        {
            if (gasConstantTextBox.Text != "" && dencityTextBox.Text != "") //protection from empty cleaning
            {
                gasNameTextBox.Text = gasConstantTextBox.Text = dencityTextBox.Text = "";
                gasListView.Items.Clear(); //clear screen 
                myGas.Clear(); //clear gas 
            }

        }

    }
}
