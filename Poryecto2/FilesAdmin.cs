using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Spire.Xls;

namespace Poryecto2
{
    class FilesAdmin
    {
        public string createFolder(string dir, string nameFolder)
        {
            System.IO.Directory.CreateDirectory(dir + "\\" + nameFolder);
            return dir + "\\" + nameFolder;
        }
        public void moveFile(string origen, string nameFile, string destino)
        {
            if (!System.IO.File.Exists(origen + "\\"))
            {
                System.IO.File.Move(origen + "\\" + nameFile, destino + "\\" + nameFile);
            }
        }

        public string consolidar(string dir, string[] listFiles){
            if (listFiles.Length != 0){
                List<string> inputFiles = new List<string>();
                foreach (string item in listFiles){
                    if (item.Contains(".xls")){
                        inputFiles.Add(item);
                    }
                }
                if (inputFiles.Count > 0){
                    Workbook newWorkbook = new Workbook();
                    newWorkbook.Worksheets.Clear();
                    Workbook tempWorkbook = new Workbook();
                    foreach (string file in inputFiles)
                    {
                        tempWorkbook.LoadFromFile(file);
                        foreach (Worksheet sheet in tempWorkbook.Worksheets)
                        {
                            newWorkbook.Worksheets.AddCopy(sheet, WorksheetCopyType.CopyAll);
                        }
                    }
                    newWorkbook.SaveToFile(dir + "\\ArchivoMaestro.xlsx", ExcelVersion.Version2013);
                }
                return dir + "\\ArchivoMaestro.xlsx";
            }
            return "";
        }
    }
}
