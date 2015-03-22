using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConvertOfficeFiles
{
    public class XlsToXlsXConverter
    {
        FileInfo[] files;
        DirectoryInfo StartingDir { get; set; }
        DirectoryInfo ArchiveDir { get; set; }

        public XlsToXlsXConverter(DirectoryInfo startingDir)
        {
            StartingDir = startingDir;
        }

        public void ConvertAll()
        {

            FindFiles(StartingDir.FullName, "*.xls");

            // only open and close Xls once
            Excel.Application msdoc = new Excel.Application();

            try
            {

                foreach (FileInfo file in files)
                {
                    if (file.ToString().ToLower().EndsWith(".xls"))
                    {
                        try
                        {
                            // convert the source file 
                            Excel.Workbook doc = msdoc.Workbooks.Open(file.FullName);
                            string newFilename = file.FullName.Replace(".xls", ".xlsx");
                            doc.SaveAs(newFilename, Excel.XlFileFormat.xlOpenXMLWorkbook, ConflictResolution: 1);
                            doc.Close();
                            file.Delete();
                        }
                        catch (System.IO.IOException e)
                        {
                            Console.WriteLine("Error reading from {0}. Message = {1}", e.Message);
                        }
                    }
                }
            }
            catch (System.IO.IOException e)
            {
                Console.WriteLine("Error reading from {0}. Message = {1}", e.Message);
            }
            finally
            {

                msdoc.Quit();
            }
        }

        void FindFiles(string sDir, string filter)
        {
            DirectoryInfo di = new DirectoryInfo(sDir);
            files = di.GetFiles(filter, SearchOption.AllDirectories);
        }

    }
} 
