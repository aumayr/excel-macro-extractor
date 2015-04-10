using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelMacroExtrator
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2 || args.Length > 3)
            {
                Console.WriteLine("Usage: excelmacroextractor.exe file targetdir [--copy-xlsm]");
                return;
            }

            ExcelExtractor excelExtr = new ExcelExtractor();
            excelExtr.ExtractMacrosFromXLSM(args[0], args[1]);

            if (args.Length == 3 && args[2] == "--copy-xlsm")
            {
                excelExtr.CopyXLSM(args[0], args[1]);
            }
        }
    }

    class ExcelExtractor
    {
        Microsoft.Office.Interop.Excel.Application _excelApp;

        public void ExtractMacrosFromXLSM(string excelFilePath, string targetPath)
        {
            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }

            _excelApp = new Microsoft.Office.Interop.Excel.Application();

            Workbook workBook = _excelApp.Workbooks.Open(excelFilePath,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing);
    
            VBProject prj;
            CodeModule code;
            string composedFile;

            prj = workBook.VBProject;
            
            // For future use:
            // prj.VBComponents.Remove(prj.VBComponents.Item(prj.VBComponents.Count-1));
            // prj.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
            // prj.VBComponents.Item(prj.VBComponents.Count).Name = "DatenspeicherAPI";
            // prj.VBComponents.Item(prj.VBComponents.Count).CodeModule.AddFromString("HALLO");

            foreach (VBComponent comp in prj.VBComponents)
            {
                code = comp.CodeModule;
                composedFile = "";
                
                for (int i = 0; i < code.CountOfLines; i++)
                {
                    composedFile += code.get_Lines(i + 1, 1) + Environment.NewLine;
                }
                   
                if (composedFile.Length > 0) {
                    Console.WriteLine("Extracting " + comp.Name + @".vba");
                    File.WriteAllText(targetPath + @"\" + comp.Name + @".vba", composedFile, Encoding.UTF8);
                }
            }

            Console.WriteLine();

            workBook.Close(false);
        }

        public void CopyXLSM(string excelFilePath, string targetPath)
        {
            string filename = Path.GetFileName(excelFilePath);
            string targetFileWithPath = targetPath + @"\" + filename;

            File.Copy(excelFilePath, targetFileWithPath, true);
        }
    }
}
