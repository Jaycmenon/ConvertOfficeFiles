using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace ConvertOfficeFiles
{
    public class DocToDocxConverter
    {
        FileInfo[] files;
        DirectoryInfo StartingDir { get; set; }
        DirectoryInfo ArchiveDir {get; set;}

        public DocToDocxConverter(DirectoryInfo startingDir)
        {
            StartingDir = startingDir;
        }

        public void ConvertAll()
        {

                FindFiles(StartingDir.FullName, "*.doc");

                // only open and close Word once
                Word.Application msdoc = new Word.Application();
                
                try
                {

                    foreach (FileInfo file in files)
                    {
                        if (file.ToString().ToLower().EndsWith(".doc"))
                        {
                            try
                            {
                                // convert the source file 
                                    var doc = msdoc.Documents.Open(file.FullName);
                                    string newFilename = file.FullName.Replace(".doc", ".docx");
                                    doc.SaveAs2(newFilename, Word.WdSaveFormat.wdFormatDocumentDefault);
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
                catch(System.IO.IOException e)
                {
                    Console.WriteLine("Error reading from {0}. Message = {1}",  e.Message);
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

