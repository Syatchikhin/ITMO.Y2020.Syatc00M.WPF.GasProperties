using System;
using Excel = Microsoft.Office.Interop.Excel;

//******************************************************************
// ��������� ��� ������� ������� �����, � 2020 �������� �.
// The program for gas properties calculation, � 2020 Syatchikhin M.
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
            //Console.WriteLine(" �������� 'Excel', �������� 30 ���.");
            Excel.Application ObjWorkExcel = new Excel.Application(); //Open Excel //2stage
            //Console.WriteLine(" �������� �����, �������� 2 ���.\n");
            //Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //������� ���� //2stage
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //������� ����
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

            for (int j = 0; j < length; j++) // �� ���� �������
            {
                created.componentName[j] = ObjWorkSheet.Cells[j + startRow, 2].Text.ToString();//��������� ����� 2 ������
                created.componentFormula[j] = ObjWorkSheet.Cells[j + startRow, 3].Text.ToString();//��������� ����� 3 ������
                created.componentData[j, 0] = double.Parse(ObjWorkSheet.Cells[j + startRow, 4].Text.ToString());//��������� ����� 4 ������
                created.componentData[j, 1] = double.Parse(ObjWorkSheet.Cells[j + startRow, 5].Text.ToString());//��������� ����� 5 ������ 
                created.componentWeight[j] = 0; // �������� ������
            }

            ObjWorkBook.Close(false, Type.Missing, Type.Missing); //������� �� ��������
            ObjWorkExcel.Quit(); // ����� �� ������
            GC.Collect(); // ������ �� ����� -- � ��� ����� �� ������������ ���� ������� !

            return created;
        }

        public static GasComposition Normalize(ref GasComposition created)
        {
            // ���������� ������� ���� � 100%
            double actualMass = 0;
            double temp;

            // ����������� �����
            for (int j = 0; j < created.size; j++) // �� ���� �������
            {
                actualMass += created.componentData[j, 1];//��������� �������� ����� �����������
            }
            for (int j = 0; j < created.size; j++) // �� ���� �������
            {
                temp = created.componentData[j, 1];//������������ �����
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
            
            for (int j = 0; j < created.size; j++) // �� ���� �������
            {
                componentVolume = created.componentData[j, 1] * 10; // Volume in liters
                componentMoleWeight = componentVolume / GasMoleVolume; // Moles amount
                created.componentWeight[j] = componentMoleWeight * created.componentData[j, 0]; //component mass
                totalWeight += created.componentWeight[j]; //full mass
            }

            created.mixtureDencity = Math.Round(totalWeight / 1000, 3, MidpointRounding.ToEven);

            for (int j = 0; j < created.size; j++) // �� ���� �������
            {
                componentMi = (created.componentWeight[j] / totalWeight) * 100;//���� ���������� �� ���� �����
                componentRi = GasConstant / created.componentData[j, 0]; //Ri
                componentMiRi = componentRi * componentMi;  //Ri*Mi
                totalMiRi += componentMiRi;
                created.mixtureR = Math.Round(totalMiRi / 100, 3, MidpointRounding.ToEven); // R rounded
            }

            return created;
        }

        //public static void OutputGasComposition(GasComposition created)
        //{
        //    //--������� ������ �� ������� ����
        //    Console.WriteLine(" ��� �����: {0}\n", created.gasName);
        //    Console.WriteLine(" #  {0,-2} {1,-20} {2,-9} {3,-8} {4}\n", " ", "���������", "�-��", "���.�", "��.%\n");
        //    for (int j = 0; j < created.size; j++)
        //    {
        //        Console.WriteLine(" �: {0,-2} {1,-20} {2,-9} {3,-8:N3} {4:N3}\n",
        //        j + 1, created.componentName[j], created.componentFormula[j],
        //        created.componentData[j, 0], created.componentData[j, 1]);
        //    }

        //}

        //public static void PrintResults(GasComposition created)
        //{
        //    //--������� ��������� �������
        //    Console.WriteLine(" ��������� (0��, 101325 ��): {0:N3} ��/�3\n", created.mixtureDencity);
        //    Console.WriteLine(" ������� ���������� �����: {0:N3} ��/(��*�)\n", created.mixtureR);
        //    Console.WriteLine(" ������� 'Enter'");
        //    Console.ReadLine();
        //}

        public static void SaveResultsToExcel(ref string savePath, ref GasComposition created)
        {
            // Save to EXCEL
           // Console.WriteLine(" �������� 'Excel', �������� 30 ���.");
            Excel.Application ObjWorkExcel = new Excel.Application(); //Open Excel
                                                                      // Console.WriteLine(" �������� �����, �������� 2 ���.\n");
            //path = MainForm.savePath;
            //Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //readonly
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(savePath, false, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //read
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            ObjWorkSheet.Cells[5, 7] = created.mixtureDencity;
            ObjWorkSheet.Cells[5, 8] = "���������(0��, 101325 ��), ��/�3";
            ObjWorkSheet.Cells[7, 7] = created.mixtureR;
            ObjWorkSheet.Cells[7, 8] = "������� ���������� �����, ��/(��*�)";

            // ObjWorkBook.Close(false, Type.Missing, Type.Missing); //������� �� ��������
            ObjWorkBook.Close(true, Type.Missing, Type.Missing); //��������� 
            //ObjWorkBook.Close(); //��������� 
            ObjWorkExcel.Quit(); // ����� �� ������

            //Console.WriteLine(" ���������� ��������� �'Excel' ����");
            //Console.WriteLine(" ������� 'Enter'");
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
        //        MessageBox.Show("����� ��� �������", "����������", MessageBoxButtons.OK, MessageBoxIcon.Information);
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


