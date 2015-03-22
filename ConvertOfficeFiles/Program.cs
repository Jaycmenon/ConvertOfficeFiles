using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace ConvertOfficeFiles
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length ==0)
            {
                Console.WriteLine("No arguments provided");
                Console.WriteLine("usage->ConvertOfficeFiles <path>");
                return;
            }
            try
            {
                string path = args[0];
                DirectoryInfo dir = new DirectoryInfo(path);
                DocToDocxConverter docs = new DocToDocxConverter(dir);
                XlsToXlsXConverter spreadsheets = new XlsToXlsXConverter(dir);
                PptToPptXConverter powerpoints = new PptToPptXConverter(dir);
                docs.ConvertAll();
                spreadsheets.ConvertAll();
                powerpoints.ConvertAll();
            }
            catch
            {
                Console.WriteLine("Wrong path/input data");
                return;
            }
            return;
        }
    }
}
