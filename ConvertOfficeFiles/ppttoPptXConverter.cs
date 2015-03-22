using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace ConvertOfficeFiles
{
    public class PptToPptXConverter
    {
        FileInfo[] files;
        DirectoryInfo StartingDir { get; set; }
        DirectoryInfo ArchiveDir { get; set; }

        public PptToPptXConverter(DirectoryInfo startingDir)
        {
            StartingDir = startingDir;
        }

        public void ConvertAll()
        {

            FindFiles(StartingDir.FullName, "*.ppt");

            // only open and close Xls once
            PowerPoint.Application msdoc = new PowerPoint.Application();

            try
            {

                foreach (FileInfo file in files)
                {
                    if (file.ToString().ToLower().EndsWith(".ppt"))
                    {
                        try
                        {
                            // convert the source file 
                            PowerPoint.Presentation doc = msdoc.Presentations.Open(file.FullName);
                            string newFilename = file.FullName.Replace(".ppt", ".pptx");
                            doc.SaveAs(newFilename, PowerPoint.PpSaveAsFileType.ppSaveAsDefault);
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
