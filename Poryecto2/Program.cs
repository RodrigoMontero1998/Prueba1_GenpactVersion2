using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Poryecto2
{
    class Program
    {
        [STAThread]
        static void Main(string[] args){
            if (args.Length == 0){
                FolderBrowserDialog fbd = new FolderBrowserDialog();
                if (fbd.ShowDialog() == DialogResult.OK){
                    string dir = fbd.SelectedPath;
                    realizarTodo(dir);
                    Console.WriteLine("Se consolido con Exito\n");
                }
            }else{
                string dir = args[0];
                realizarTodo(dir);
                Console.WriteLine("Se consolido con Exito\n");
            }
            Console.WriteLine("El programa termino\n");
        }

        static void realizarTodo(string dir){
            FilesAdmin f = new FilesAdmin();
            string aux = f.consolidar(dir, Directory.GetFiles(dir));
            if(aux != "") { 
                foreach (string Item in Directory.GetFiles(dir)){
                    FileInfo fi = new FileInfo(Item);
                    moveFiles(dir,fi.Name);
                }
            }
            else{
                Console.WriteLine("No se ejecuto con exito\n");
            }
        }

        static void moveFiles(string dir,string nameFile)
        {
            if(nameFile != "ArchivoMaestro.xlsx")
            {
                FilesAdmin f = new FilesAdmin();
                if (nameFile.Contains(".xls"))
                {
                    string destino = f.createFolder(dir, "Procesado");
                    f.moveFile(dir, nameFile, destino);

                }
                else
                {
                    string destino = f.createFolder(dir, "No aplicable");
                    f.moveFile(dir, nameFile, destino);
                }
            }
           
        }
    }

    
}
