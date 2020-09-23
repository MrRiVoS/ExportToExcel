using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;

namespace GetResultFromUTCOMP
{
    class Program
    {
        static void Main(string[] args)
        {
            int nprints; //= 0;
            DataSet time, oil_rate, gas_rate;
            string filename = @"C:\Users\ivens\Desktop\TEST.HIS";

            Console.WriteLine("\n\n Getting number of prints...");
            nprints = ValuesFromFile.GetNumberOfPrints(filename);

            time = new DataSet(nprints);
            oil_rate = new DataSet(nprints);
            gas_rate = new DataSet(nprints);

            Console.WriteLine(" Getting surface production...");
            ValuesFromFile.GetData(filename, nprints, time, oil_rate, gas_rate);

            Console.WriteLine(" Exportign file to Excel...");
            ValuesFromFile.ExportToExcel();
        }
    }
}
