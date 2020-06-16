using System;
using Excel = Microsoft.Office.Interop.Excel;

//******************************************************************
// Программа для расчета свойств газов, © 2020 Сятчихин М.
// The program for gas properties calculation, © 2020 Syatchikhin M.
// *****************************************************************

namespace WpfApp1
{

    public class GasComposition
    {
        public string gasName;
        public int size;
        public string[] componentName;
        public string[] componentFormula;
        public double[,] componentData;
        public double[] componentWeight;

        public double mixtureDencity;
        public double mixtureR;


        public static GasComposition ReadExcelFile(ref string path, ref GasComposition created)
        {
            //Read Excel
            //Console.WriteLine(" Открытие 'Excel', примерно 30 сек.");
            Excel.Application ObjWorkExcel = new Excel.Application(); //Open Excel //2stage
            //Console.WriteLine(" Открытие файла, примерно 2 сек.\n");
            //Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл //2stage
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //Get 1 sheet
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 cell
                                                                                                //-------------------------------------
            int lastRow = (int)lastCell.Row;
            //-------------------------------------
            int startRow = 5;
            string tempGasName = ObjWorkSheet.Cells[2, 1].Text.ToString();//Read gas name
            int length = lastRow - startRow + 1;
            //---------------------------------------
            //GasComposition created = new GasComposition();
            created.size = length;
            created.gasName = tempGasName;
            created.componentName = new string[length];
            created.componentFormula = new string[length];
            created.componentData = new double[length, 2];
            created.componentWeight = new double[length]; // array for calc purposes

            for (int j = 0; j < length; j++) // по всем строкам
            {
                created.componentName[j] = ObjWorkSheet.Cells[j + startRow, 2].Text.ToString();//считываем текст 2 строки
                created.componentFormula[j] = ObjWorkSheet.Cells[j + startRow, 3].Text.ToString();//считываем текст 3 строки
                created.componentData[j, 0] = double.Parse(ObjWorkSheet.Cells[j + startRow, 4].Text.ToString());//считываем текст 4 строки
                created.componentData[j, 1] = double.Parse(ObjWorkSheet.Cells[j + startRow, 5].Text.ToString());//считываем текст 5 строки 
                created.componentWeight[j] = 0; // Занулить массив
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkExcel.Quit(); // выйти из экселя
            GC.Collect(); // убрать за собой -- в том числе не используемые явно объекты !

            return created;
        }

        public static GasComposition Normalize(ref GasComposition created)
        {
            // Приведение состава газа к 100%
            double actualMass = 0;
            double temp;

            // фактическая масса
            for (int j = 0; j < created.size; j++) // по всем строкам
            {
                actualMass += created.componentData[j, 1];//считываем значения массы компонентов
            }
            for (int j = 0; j < created.size; j++) // по всем строкам
            {
                temp = created.componentData[j, 1];//корректируем массу
                created.componentData[j, 1] = (temp * 100) / actualMass;
            }
            return created;
        }

        public static GasComposition CalculateProperties(ref GasComposition created)
        {
            const double GasConstant = 8314.462618;
            const double GasMoleVolume = 22.41396954;
            double totalMiRi = 0;
            double totalWeight = 0;
            double componentVolume;
            double componentMoleWeight;
            double componentMi;
            double componentRi;
            double componentMiRi;
            
            for (int j = 0; j < created.size; j++) // по всем строкам
            {
                componentVolume = created.componentData[j, 1] * 10; // Volume in liters
                componentMoleWeight = componentVolume / GasMoleVolume; // Moles amount
                created.componentWeight[j] = componentMoleWeight * created.componentData[j, 0]; //component mass
                totalWeight += created.componentWeight[j]; //full mass
            }

            created.mixtureDencity = Math.Round(totalWeight / 1000, 3, MidpointRounding.ToEven);

            for (int j = 0; j < created.size; j++) // по всем строкам
            {
                componentMi = (created.componentWeight[j] / totalWeight) * 100;//доля компонента от всей массы
                componentRi = GasConstant / created.componentData[j, 0]; //Ri
                componentMiRi = componentRi * componentMi;  //Ri*Mi
                totalMiRi += componentMiRi;
                created.mixtureR = Math.Round(totalMiRi / 100, 3, MidpointRounding.ToEven); // R rounded
            }

            return created;
        }

        //public static void OutputGasComposition(GasComposition created)
        //{
        //    //--выводим данные по составу газа
        //    Console.WriteLine(" Имя смеси: {0}\n", created.gasName);
        //    Console.WriteLine(" #  {0,-2} {1,-20} {2,-9} {3,-8} {4}\n", " ", "Компонент", "Ф-ла", "Мол.м", "Об.%\n");
        //    for (int j = 0; j < created.size; j++)
        //    {
        //        Console.WriteLine(" №: {0,-2} {1,-20} {2,-9} {3,-8:N3} {4:N3}\n",
        //        j + 1, created.componentName[j], created.componentFormula[j],
        //        created.componentData[j, 0], created.componentData[j, 1]);
        //    }

        //}

        //public static void PrintResults(GasComposition created)
        //{
        //    //--выводим результат расчета
        //    Console.WriteLine(" Плотность (0°С, 101325 Па): {0:N3} кг/м3\n", created.mixtureDencity);
        //    Console.WriteLine(" Газовая постоянная смеси: {0:N3} Дж/(кг*К)\n", created.mixtureR);
        //    Console.WriteLine(" Нажмите 'Enter'");
        //    Console.ReadLine();
        //}

        public static void SaveResultsToExcel(ref string savePath, ref GasComposition created)
        {
            // Save to EXCEL
           // Console.WriteLine(" Открытие 'Excel', примерно 30 сек.");
            Excel.Application ObjWorkExcel = new Excel.Application(); //Open Excel
                                                                      // Console.WriteLine(" Открытие файла, примерно 2 сек.\n");
            //path = MainForm.savePath;
            //Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //readonly
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(savePath, false, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //read
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            ObjWorkSheet.Cells[5, 7] = created.mixtureDencity;
            ObjWorkSheet.Cells[5, 8] = "Плотность(0°С, 101325 Па), кг/м3";
            ObjWorkSheet.Cells[7, 7] = created.mixtureR;
            ObjWorkSheet.Cells[7, 8] = "Газовая постоянная смеси, Дж/(кг*К)";

            // ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
            ObjWorkBook.Close(true, Type.Missing, Type.Missing); //сохранить 
            //ObjWorkBook.Close(); //сохранить 
            ObjWorkExcel.Quit(); // выйти из экселя

            //Console.WriteLine(" Результаты сохранены в'Excel' файл");
            //Console.WriteLine(" Нажмите 'Enter'");
            //Console.ReadLine();
        }

        //public void CleanScreen()
        //{
        //    if (gasConstantTextBox.Text != "" && dencityTextBox.Text != "") //protection from empty cleaning
        //    {
        //        gasNameTextBox.Text = gasConstantTextBox.Text = dencityTextBox.Text = "";
        //        gasListView.Items.Clear(); //clear screen 
        //        myGas.Clear(); //clear gas 
        //    }
        //    else
        //    {
        //        MessageBox.Show("Форма уже очищена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //    }
        //}

        //public void CleanScreenNoMessage()
        //{
        //    if (gasConstantTextBox.Text != "" && dencityTextBox.Text != "") //protection from empty cleaning
        //    {
        //        gasNameTextBox.Text = gasConstantTextBox.Text = dencityTextBox.Text = "";
        //        gasListView.Items.Clear(); //clear screen 
        //        myGas.Clear(); //clear gas 
        //    }

        //}




    }

}

//ReadFile() {}

//SaveResltsToFile() {}


