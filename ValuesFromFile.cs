using DocumentFormat.OpenXml.Office2010.Excel.Drawing;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Security.Cryptography;
using Excel = Microsoft.Office.Interop.Excel;

public class ValuesFromFile
{
	public static int GetNumberOfPrints(string filename)
	{
        using (StreamReader fs = File.OpenText(filename))
        {
            bool keep_reading = true;
            while (keep_reading)
            {
                string line = fs.ReadLine();

                if (line.Contains("PRINTS"))
                {
                    string praver = fs.ReadLine();
                    bool worked = int.TryParse(praver, out int nprints);
                    if (worked)
                    {
                        return nprints;
                    }
                    else
                    {
                        Console.WriteLine("Couldn't parse number of prints on file: " + filename);
                        return 0;
                    }
                }
            }
            Console.WriteLine("Couldn't find number of prints on file: " + filename);
            return 0;
        }
	}

    public static void GetData(string filename, int nprints, DataSet var1, DataSet var2, DataSet var3)
    {
        using (StreamReader fs = File.OpenText(filename))
        {
            bool keep_reading = true;
            while (keep_reading)
            {
                string line = fs.ReadLine();

                if (line.Contains("TOTAL SURFACE PRODUCTION"))
                {
                    string line2 = fs.ReadLine();
                    string temp;

                    for (int i = 0; i < nprints; i++)
                    {
                        // Read line from file
                        temp = fs.ReadLine();

                        // Split line in white spaces
                        string[] stnumbers = temp.Split(" ");

                        // Parse splited strings in a double array
                        int count = 0;
                        double[] numbers = new double[5];
                        foreach (var item in stnumbers)
                        {
                            if(!string.IsNullOrEmpty(item))
                            {
                                bool ok = double.TryParse(item, out numbers[count]);
                                count += 1;
                            }
                        }

                        // Pass values to DataSet objects
                        var1.SetData(i, numbers[0]);
                        var2.SetData(i, numbers[3]);
                        var3.SetData(i, numbers[4]);
                    }
                    break;
                }
            }
        }
    }

    public static void ExportToExcel(int ndata, DataSet var1, DataSet var2, DataSet var3)
    {
        // Create Excel application (it's like opening the program)
        Excel.Application excelApp = new Excel.Application();

        // Add a workbook (it's like creating a new empty file with one empty active worksheet)
        excelApp.Workbooks.Add();

        // Create a local worksheet and gets the reference of the active worksheet in Excel (excelApp)
        Excel._Worksheet workSheet = (Excel._Worksheet)excelApp.ActiveSheet;

        //----------------------------------------
        // Filling the worksheet with data
        //----------------------------------------

        // Setting "table name"
        Excel.Range headerRange = excelApp.get_Range("A1", "C1");
        headerRange.Merge();
        workSheet.Cells[1, "A"] = "<Case name>";
        headerRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
        headerRange.Font.Bold = true;
        headerRange.Interior.Color = System.Drawing.Color.Yellow;

        // Writting headers
        workSheet.Cells[2, "A"] = "Time (days)";
        workSheet.Cells[2, "B"] = "Oil rate (STB/d)";
        workSheet.Cells[2, "C"] = "Gas rate (STB/d)";

        // Writting data
        int counter = 2;
        for (int i = 0; i < ndata; i++)
        {
            counter ++;

            workSheet.Cells[counter, "A"] = var1.GetDataItem(i);
            workSheet.Cells[counter, "B"] = var2.GetDataItem(i);
            workSheet.Cells[counter, "C"] = var3.GetDataItem(i);
        }

        // Formating data range
        string beg_range = "A3";
        int endRow = ndata + 2;
        string end_range = "C" + endRow;
        Excel.Range dataRange = excelApp.get_Range(beg_range, end_range);
        //dataRange.Cells.NumberFormat = "#######";

        // Formating the data
        //workSheet.Range["A2"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatTable3);

        //----------------------------------------
        // Creating the chart
        //----------------------------------------

        // Creating the objects
        var charts = workSheet.ChartObjects() as Excel.ChartObjects;
        var chartObject = charts.Add(100, 100, 500, 200);
        var oilChart = chartObject.Chart;
        Excel.Chart gasChart = new Excel.Chart();
        end_range = "A" + endRow;
        Excel.Range timeRange = excelApp.get_Range("A3", end_range);
        end_range = "B" + endRow;
        Excel.Range oilRange = excelApp.get_Range("B3", end_range);

        // Filling chart with data
        oilChart.SetSourceData(oilRange);
        //oilChart.

        // Setting chart type
        oilChart.ChartType = Excel.XlChartType.xlXYScatterSmoothNoMarkers;


        // Another approach
        //Excel.Series oilSeries = null;
        //Excel.Shape _Shape = workSheet.Shapes.AddChart2();

        //oilSeries.XValues = dataRange;
        //_Shape.Chart.SeriesCollection. = oilSeries;

        // Saving the file
        workSheet.SaveAs($@"{Environment.CurrentDirectory}\outfile.xlsx");

        excelApp.Quit();

    }
}
