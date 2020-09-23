using System;
using System.IO;
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

    public static void ExportToExcel()
    {
        // Create Excel application (it's like opening the program)
        Excel.Application excelApp = new Excel.Application();

        // Add a workbook (it's like creating a new empty file with one empty active worksheet)
        excelApp.Workbooks.Add();

        // Create a local worksheet and gets the reference of the active worksheet in Excel (excelApp)
        Excel._Worksheet workSheet = (Excel._Worksheet)excelApp.ActiveSheet;

        // Filling the worksheet with data
        workSheet.Cells[1, "A"] = "Modelo";
        workSheet.Cells[1, "B"] = "Fabricante";
        workSheet.Cells[1, "C"] = "Ano";

        workSheet.Cells[2, "A"] = "Clio";
        workSheet.Cells[2, "B"] = "Renault";
        workSheet.Cells[2, "C"] = 2011;

        workSheet.Cells[3, "A"] = "Palio";
        workSheet.Cells[3, "B"] = "Fiat";
        workSheet.Cells[3, "C"] = 2012;

        workSheet.Cells[4, "A"] = "Sentra";
        workSheet.Cells[4, "B"] = "Nissan";
        workSheet.Cells[4, "C"] = 2016;

        // Formatign the data
        workSheet.Range["A1"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatTable1);

        // Saving the file
        workSheet.SaveAs($@"{Environment.CurrentDirectory}\outfile.xlsx");

        excelApp.Quit();
    }
}
