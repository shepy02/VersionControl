using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ExcelGeneration
{
    public partial class Form1 : Form
    {
        RealEstateEntities context = new RealEstateEntities();
        List<Flat> Flats;
        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;

        private readonly string[] headers = new string[] {
            "Kód",
            "Eladó",
            "Oldal",
            "Kerület",
            "Lift",
            "Szobák száma",
            "Alapterület (m2)",
            "Ár (mFt)",
            "Négyzetméter ár (Ft/m2)"};

        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();
        }

        private void LoadData()
        {
            Flats = context.Flats.ToList();
        }

        private void CreateExcel() 
        {
            try
            {
                xlApp = new Excel.Application();

                xlWB = xlApp.Workbooks.Add(Missing.Value);

                xlSheet = xlWB.ActiveSheet;

                CreateTable();
                FormatTable();

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        private void CreateTable()
        {
            for (int i = 0, j = 1; i < headers.Length; i++, j++)
            {
                xlSheet.Cells[1, j] = headers[i];
            }

            object[,] values = new object[Flats.Count, headers.Length];
            int k = 0;
            foreach (Flat flat in Flats)
            {
                values[k, 0] = flat.Code;
                values[k, 1] = flat.Vendor;
                values[k, 2] = flat.Side;
                values[k, 3] = flat.District;
                values[k, 4] = flat.Elevator ? "Van" : "Nincs";
                values[k, 5] = flat.NumberOfRooms;
                values[k, 6] = flat.FloorArea;
                values[k, 7] = flat.Price;
                values[k, 8] = "=" + GetCell(2+k, 8) + "*1000000/" + GetCell(2+k, 7);
                k++;
            }

            xlSheet.get_Range(
             GetCell(2, 1),
             GetCell(1 + values.GetLength(0), values.GetLength(1))).Value2 = values;
        }

        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }

        private void FormatTable()
        {
            int lastRowID = xlSheet.UsedRange.Rows.Count;

            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range valueRange = xlSheet.get_Range(GetCell(2, 1), GetCell(lastRowID, headers.Length));
            valueRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range firstColumnValueRange = xlSheet.get_Range(GetCell(2, 1), GetCell(lastRowID, 1));
            firstColumnValueRange.Font.Bold = true;
            firstColumnValueRange.Interior.Color = Color.LightYellow;

            Excel.Range lastColumnValueRange = xlSheet.get_Range(GetCell(2, headers.Length), GetCell(lastRowID, headers.Length));
            lastColumnValueRange.Interior.Color = Color.LightGreen;
            lastColumnValueRange.NumberFormat = "#,#.00";
        }
    }
}
