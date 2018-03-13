using console_excel_generator.Excel;
using console_excel_generator.Model;
using Excel.Contracts;
using Excel.Factory;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;

namespace console_excel_generator
{
    class Program
    {
        static void Main(string[] args)
        {
            //Pacote Nuget a ser instalado
            //Install-Package EPPlus -Version 4.1.1

            var r = new Random();
            var data = new List<Pai>();
            //data.Add(new Pai(profundidade: 7, largura: 100,nivel:1,r:r));
            //data.Add(new Pai(profundidade: 5, largura: 100,nivel:1,r:r));
            //data.Add(new Pai(profundidade: 2, largura: 100,nivel:1,r:r));
            //data.Add(new Pai(profundidade: 3, largura: 100,nivel:1,r:r));
            data.Add(new Pai(profundidade: 10, largura: 1000,nivel:1,r:r));
            //data.Add(new Pai(profundidade: 3, largura: 1000,nivel:1,r:r));
            var maiorProfundidade = data.Max(e=>e.profundidade);

            var a = new ExcelGenerator(@"D:\20180308 MUTANT - EXCEL\console-excel-generator\Planilhas\", data,maiorProfundidade).ExportComplete();

        }

      





    }
}
