using console_excel_generator.Model;
using Excel.Contracts;
using Excel.Factory;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace console_excel_generator.Excel
{
    public class ExcelGenerator
    {
        private ICellFactory CellFactory = new CellFactory();
        private IStyleFactory StyleFactory = new StyleFactory();

        private ExcelWorksheet ws;
        private int headerTextRow;
        private int _Column;
        private string path;
        private int _Row;
        private int fator;

        List<Pai> data { get; set; }

        public ExcelGenerator(string path, List<Pai>  data,int profundidade)
        {
            this.path = path;
            this.data = data;
            this.fator = 255/(profundidade+10);
        }

        public string StartExport()
        {
            string excelName = "Lançamentos contábeis - " + DateTime.Now.ToString("yyyyMMdd-hhmmss") + ".xlsx";

            System.IO.FileInfo newFile = new System.IO.FileInfo(path + excelName);
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new System.IO.FileInfo(path + excelName);
            }

            //Create the workbook
            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheet
            ws = pck.Workbook.Worksheets.Add("Lançamentos");

            CreateExcelHeader(ws);
            data.ForEach(e => CreateExcelLine(e));


            ws.Cells[ws.Dimension.Address].AutoFitColumns();
            ws.View.FreezePanes(2, 1);

            pck.Save();
            return excelName;
        }


        private void CreateExcelHeader(ExcelWorksheet ws)
        {
            //string colorHeader = "#233445";
            ws.OutLineSummaryRight = false;
            ws.OutLineSummaryBelow = false;
            ws.Column(1).Width = 40;

            var headerTextRow = 2;

            List<string> headers = null;

            headers = new List<string> { "Agrupamento Centro de Custos", "Código Centro de Custo", "Descrição Centro de Custo", "Ramo", "Código Item Contábil", "Descrição Item Contábil", "Código Classe de Valor", "Descrição Classe de Valor", "Agrupamento Plano de Contas", "Código Plano de Contas", "Descrição Plano de Contas", "Data", "Valor", "Histórico" };


            for (int i = 0; i < headers.Count; i++)
            {
                StyleFactory.CreateStyleHorizontalAlignment(ExcelHorizontalAlignment.Center).ApllyStyle(ws.Cells[headerTextRow, (i + 1)]);
                StyleFactory.CreateStylePatternType(ExcelFillStyle.Solid).ApllyStyle(ws.Cells[headerTextRow, (i + 1)]);
                StyleFactory.CreateStyleBackgroundColor("#000000").ApllyStyle(ws.Cells[headerTextRow, (i + 1)]);
                StyleFactory.CreateStyleFontBold().ApllyStyle(ws.Cells[headerTextRow, (i + 1)]);
                StyleFactory.CreateStyleFontColor("#FFFFFF").ApllyStyle(ws.Cells[headerTextRow, (i + 1)]);
                StyleFactory.CreateStyleBorderRight(ExcelBorderStyle.Thin).ApllyStyle(ws.Cells[headerTextRow, (i + 1)]);

                CellFactory.CreateCellText(headers[i]).ApllyCell(ws.Cells[headerTextRow, (i + 1)]);
            }

            ws.Cells[headerTextRow, 1, headerTextRow, headers.Count].AutoFilter = true;
        }
        

        private void CreateExcelLine(Pai row)
        {

            headerTextRow++;
            _Column = 1;

            //CellFactory.CreateCellText(GetAgrupamentoCC(row.valo.ToString())).ApllyCell(ws.Cells[headerTextRow, _Column++]);
            CellFactory.CreateCellMoney(row.valor1).ApllyCell(ws.Cells[headerTextRow, _Column++]);
            CellFactory.CreateCellMoney(row.valor2).ApllyCell(ws.Cells[headerTextRow, _Column++]);
            CellFactory.CreateCellMoney(row.valor3).ApllyCell(ws.Cells[headerTextRow, _Column++]);
            CellFactory.CreateCellMoney(row.valor4).ApllyCell(ws.Cells[headerTextRow, _Column++]);
            CellFactory.CreateCellMoney(row.valor5).ApllyCell(ws.Cells[headerTextRow, _Column++]);
            CellFactory.CreateCellMoney(row.valor6).ApllyCell(ws.Cells[headerTextRow, _Column++]);
            CellFactory.CreateCellMoney(row.valor7).ApllyCell(ws.Cells[headerTextRow, _Column++]);
            CellFactory.CreateCellMoney(row.valor8).ApllyCell(ws.Cells[headerTextRow, _Column++]);

            if (row.filhos!=null && row.filhos.Count > 0) {
                row.filhos.ForEach(e => CreateExcelLine(e));
            }
        }


        private List<string> GetPeriods(string type)
        {
            var result = new List<string>();

            switch (type)
            {
                case "m":
                    for (int i = 1; i <= 13; i++)
                    {
                        string period = string.Empty;
                        if (i < 13)
                            period = (i < 10 ? string.Format("0{0}", i) : i.ToString());
                        else
                            period = "sum";
                        result.Add(period);
                    };
                    break;
                case "q":
                    for (int i = 1; i <= 5; i++)
                    {
                        if (i < 5)
                            result.Add(string.Format("Q{0}", i));
                        else
                            result.Add("sum");
                    }
                    break;
                case "a":
                    result.Add("2017");
                    break;
                default:
                    break;
            }
            return result;
        }


        public string ExportComplete()
        {
            string year = "2018";
            string type ="q";
            string company = "10";
            int visionId = 1;
            var periods = GetPeriods(type);

            string excelName = string.Format("Fluxo DRE - {0}.xlsx", DateTime.Now.ToString("yyyyMMdd-hhmmss"));

            //Check file existence
            System.IO.FileInfo newFile = new System.IO.FileInfo(path + excelName);
            if (newFile.Exists)
            {
                newFile.Delete();  // ensures we create a new workbook
                newFile = new System.IO.FileInfo(path + excelName);
            }

            //Create the workbook
            ExcelPackage pck = new ExcelPackage(newFile);

            //Add the Content sheet
            ws = pck.Workbook.Worksheets.Add("Modelo");

            //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-//
            //          Create headers      //
            //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-//
            string[] headers = GetHeadersByType(type);

            ws.OutLineSummaryRight = false;
            ws.OutLineSummaryBelow = false;


            //Create first line of header
            ws.Column(1).Width = 40;

            //Create second line of header
            _Row = 1;
            _Column = 1;

            // Linha de cabeçalho 1
            StyleFactory.CreateStyleHorizontalAlignment(ExcelHorizontalAlignment.Center).ApllyStyle(ws.Cells[_Row, _Column]);
            StyleFactory.CreateStylePatternType(ExcelFillStyle.Solid).ApllyStyle(ws.Cells[_Row, _Column]);
            StyleFactory.CreateStyleBackgroundColor("#0067a1").ApllyStyle(ws.Cells[_Row, _Column]);
            StyleFactory.CreateStyleFontBold().ApllyStyle(ws.Cells[_Row, _Column]);
            StyleFactory.CreateStyleFontColor("#FFFFFF").ApllyStyle(ws.Cells[_Row, _Column]);
            CellFactory.CreateCellText("").ApllyCell(ws.Cells[_Row, _Column]);

            GetPeriods(type).ForEach(period =>
            {
                _Column++;

                ws.Cells[_Row, _Column, _Row, (_Column + headers.Length - 1)].Merge = true;

                StyleFactory.CreateStyleHorizontalAlignment(ExcelHorizontalAlignment.Center).ApllyStyle(ws.Cells[_Row, _Column]);
                StyleFactory.CreateStylePatternType(ExcelFillStyle.Solid).ApllyStyle(ws.Cells[_Row, _Column]);
                StyleFactory.CreateStyleBackgroundColor("#0067a1").ApllyStyle(ws.Cells[_Row, _Column]);
                StyleFactory.CreateStyleFontBold().ApllyStyle(ws.Cells[_Row, _Column]);
                StyleFactory.CreateStyleFontColor("#FFFFFF").ApllyStyle(ws.Cells[_Row, _Column]);
                StyleFactory.CreateStyleBorderRight(ExcelBorderStyle.Thin).ApllyStyle(ws.Cells[_Row, _Column, _Row, (_Column + headers.Length - 1)]);

                CellFactory.CreateCellText("Não sei").ApllyCell(ws.Cells[_Row, _Column, _Row, (_Column + headers.Length - 1)]);
                _Column += headers.Length - 1;
            });

            _Row++;
            _Column = 1;

            // Linha de cabeçalho 2
            StyleFactory.CreateStyleHorizontalAlignment(ExcelHorizontalAlignment.Center).ApllyStyle(ws.Cells[_Row, _Column]);
            StyleFactory.CreateStylePatternType(ExcelFillStyle.Solid).ApllyStyle(ws.Cells[_Row, _Column]);
            StyleFactory.CreateStyleBackgroundColor("#0067a1").ApllyStyle(ws.Cells[_Row, _Column]);
            StyleFactory.CreateStyleFontBold().ApllyStyle(ws.Cells[_Row, _Column]);
            StyleFactory.CreateStyleFontColor("#FFFFFF").ApllyStyle(ws.Cells[_Row, _Column]);
            CellFactory.CreateCellText("").ApllyCell(ws.Cells[_Row, _Column]);


            for (int i = 1; i <= periods.Count; i++)
            {
                for (int j = 0; j < headers.Length; j++)
                {
                    _Column++;

                    StyleFactory.CreateStyleHorizontalAlignment(ExcelHorizontalAlignment.Center).ApllyStyle(ws.Cells[_Row, _Column]);
                    StyleFactory.CreateStylePatternType(ExcelFillStyle.Solid).ApllyStyle(ws.Cells[_Row, _Column]);
                    StyleFactory.CreateStyleBackgroundColor("#0067a1").ApllyStyle(ws.Cells[_Row, _Column]);
                    StyleFactory.CreateStyleFontBold().ApllyStyle(ws.Cells[_Row, _Column]);
                    StyleFactory.CreateStyleFontColor("#FFFFFF").ApllyStyle(ws.Cells[_Row, _Column]);

                    if (headers.Length - 1 == j)
                        StyleFactory.CreateStyleBorderRight(ExcelBorderStyle.Thin).ApllyStyle(ws.Cells[_Row, _Column]);

                    CellFactory.CreateCellText(headers[j]).ApllyCell(ws.Cells[_Row, _Column]);
                }
            }

            switch (type)
            {
                case "m":
                case "q":
                    StartCreateLines(year, company, visionId, type, periods, data);
                    break;
                case "a":
                    //StarCreateLinesForYearReport(year, company, visionId, type, periods);
                    break;
                default:
                    break;
            }

            //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-//
            //          Write Data          //
            //=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-//

            pck.Save();
            return excelName;
        }


        private string[] GetHeadersByType(string type)
        {
            var result = new List<string>();

            switch (type)
            {
                case "m":
                case "q":
                    result = new List<string>() { "Orçado", "Forecast", "Realizado", "Orçado x Realizado", "Forecast x Realizado", "OxR%", "FxR%" };
                    break;
                case "a":
                    result = new List<string>() { "JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ", "TOTAL" };
                    break;
            }

            return result.ToArray();
        }


        private void StartCreateLines(string year, string company, int visionId, string type, List<string> periods, List<Pai> data)
        {
            //string collection = "budgetapplines";
            //List<LinhaDataDRE> structureLines = GetLines(year, company, visionId, 0, collection).OrderBy(x => x.data.OrdemExibicao).ToList();

            //data.ForEach(e=> {
            //    CreateLineQuick(e, type, periods,);
                
            //    });


            var pointerRow = 1;
            var ret = 0;
            foreach (var level in data)
            {
                // subLevel.data.Children = line.data.Children.Where(x => x.data.IdPai.Contains(string.Format("{0}_{1}", subLevel.data.IdPai, subLevel.data.Cod))).ToList();
                ret = CreateLineQuick(level, type, periods, pointerRow++, ret);
            }


        }


        private void CreateLine(Pai line, string type, List<string> periods)
        {
            _Row++;
            _Column = 1;

            ws.Row(_Row).Collapsed = true;
            ws.Row(_Row).OutlineLevel = line.nivel;

            //if (line.data.Nivel > 7)
            //    line.data.Nivel = 7;

            CellFactory.CreateCellText(line.desc).ApllyCell(ws.Cells[_Row, _Column]);

            StyleFactory.CreateStyleFontColor("#000").ApllyStyle(ws.Cells[_Row, 1]);
            StyleFactory.CreateStylePatternType(ExcelFillStyle.Solid).ApllyStyle(ws.Cells[_Row, 1]);
            StyleFactory.CreateStyleBackgroundColor(GetBackgroundColor(line.nivel)).ApllyStyle(ws.Cells[_Row, 1]);
            StyleFactory.CreateStyleIndent(((line.nivel * 3) < 0 ? 0 : (line.nivel * 3))).ApllyStyle(ws.Cells[_Row, 1]);


            periods.ForEach(period =>
            {
                _Column++;
                ws.Column(_Column).Width = 15;

             //   IDictionary<string, object> values = arrValues[period] as IDictionary<string, object>;

                decimal orcado = line.valor1;
                decimal forcast = line.valor2;
                decimal realizado = line.valor3;
                decimal orcadoVersusRealizado = realizado - orcado;
                decimal forcastVersusRealizado = realizado - forcast;
                decimal percentualOrcadoVersusRealizado = orcado != 0 ? ((realizado - orcado) / orcado) : 0;
                decimal percentualForcastVersusRealizado = forcast != 0 ? ((realizado - forcast) / forcast) : 0;

                int resultsQuantity = 7;
                int column = _Column + resultsQuantity;

                StyleFactory.CreateStylePatternType(ExcelFillStyle.Solid).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                StyleFactory.CreateStyleBackgroundColor(GetBackgroundColor(line.nivel)).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);

                StyleFactory.CreateStyleFontColor((orcado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                CellFactory.CreateCellMoney(orcado).ApllyCell(ws.Cells[_Row, _Column]);

                StyleFactory.CreateStyleFontColor((orcado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                CellFactory.CreateCellMoney(forcast).ApllyCell(ws.Cells[_Row, ++_Column]);
                ws.Column(_Column).Width = 15;

                StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                CellFactory.CreateCellMoney(realizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                ws.Column(_Column).Width = 15;

                // Calculos
                StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                CellFactory.CreateCellMoney(orcadoVersusRealizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                ws.Column(_Column).Width = 15;

                StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                CellFactory.CreateCellMoney(forcastVersusRealizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                ws.Column(_Column).Width = 15;

                StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                CellFactory.CreateCellPercentage(percentualOrcadoVersusRealizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                ws.Column(_Column).Width = 15;

                StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                CellFactory.CreateCellPercentage(percentualForcastVersusRealizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                ws.Column(_Column).Width = 15;

                StyleFactory.CreateStyleBorderRight(ExcelBorderStyle.Thin).ApllyStyle(ws.Cells[_Row, _Column]);

            });

            var subLevels = line.filhos;
            //var subLevels = line.data.Children;

            if (subLevels != null)
            {
                foreach (var subLevel in subLevels)
                {
                   // subLevel.data.Children = line.data.Children.Where(x => x.data.IdPai.Contains(string.Format("{0}_{1}", subLevel.data.IdPai, subLevel.data.Cod))).ToList();
                    CreateLine(subLevel, type, periods);
                }
            }
        }



        private int CreateLineQuick(Pai line, string type, List<string> periods, int row, int firstRow)
        {
            int retorno = 0;
            _Row++;
            _Column = 1;

            ws.Row(_Row).Collapsed = true;
            ws.Row(_Row).OutlineLevel = line.nivel;

            //if (line.data.Nivel > 7)
            //    line.data.Nivel = 7;

            CellFactory.CreateCellText(line.desc).ApllyCell(ws.Cells[_Row, _Column]);

            StyleFactory.CreateStyleFontColor("#000").ApllyStyle(ws.Cells[_Row, 1]);
            StyleFactory.CreateStylePatternType(ExcelFillStyle.Solid).ApllyStyle(ws.Cells[_Row, 1]);
            StyleFactory.CreateStyleBackgroundColor(GetBackgroundColor(line.nivel)).ApllyStyle(ws.Cells[_Row, 1]);
            StyleFactory.CreateStyleIndent(((line.nivel * 3) < 0 ? 0 : (line.nivel * 3))).ApllyStyle(ws.Cells[_Row, 1]);


            periods.ForEach(period =>
            {
                _Column++;
                ws.Column(_Column).Width = 15;

                //   IDictionary<string, object> values = arrValues[period] as IDictionary<string, object>;

                decimal orcado = line.valor1;
                decimal forcast = line.valor2;
                decimal realizado = line.valor3;
                decimal orcadoVersusRealizado = realizado - orcado;
                decimal forcastVersusRealizado = realizado - forcast;
                decimal percentualOrcadoVersusRealizado = orcado != 0 ? ((realizado - orcado) / orcado) : 0;
                decimal percentualForcastVersusRealizado = forcast != 0 ? ((realizado - forcast) / forcast) : 0;
                int resultsQuantity = 7;
                int column = 0;

                if (row == 1)
                {
                    if (_Column == 2)
                    {
                        column = _Column + resultsQuantity;

                        StyleFactory.CreateStylePatternType(ExcelFillStyle.Solid).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                        StyleFactory.CreateStyleBackgroundColor(GetBackgroundColor(line.nivel)).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);

                        StyleFactory.CreateStyleFontColor((orcado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                        CellFactory.CreateCellMoney(orcado).ApllyCell(ws.Cells[_Row, _Column]);

                        StyleFactory.CreateStyleFontColor((orcado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                        CellFactory.CreateCellMoney(forcast).ApllyCell(ws.Cells[_Row, ++_Column]);
                        ws.Column(_Column).Width = 15;

                        StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                        CellFactory.CreateCellMoney(realizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                        ws.Column(_Column).Width = 15;

                        // Calculos
                        StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                        CellFactory.CreateCellMoney(orcadoVersusRealizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                        ws.Column(_Column).Width = 15;

                        StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                        CellFactory.CreateCellMoney(forcastVersusRealizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                        ws.Column(_Column).Width = 15;

                        StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                        CellFactory.CreateCellPercentage(percentualOrcadoVersusRealizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                        ws.Column(_Column).Width = 15;

                        StyleFactory.CreateStyleFontColor((realizado < 0 ? "#FF0000" : "#000")).ApllyStyle(ws.Cells[_Row, _Column, _Row, column]);
                        CellFactory.CreateCellPercentage(percentualForcastVersusRealizado).ApllyCell(ws.Cells[_Row, ++_Column]);
                        ws.Column(_Column).Width = 15;

                        StyleFactory.CreateStyleBorderRight(ExcelBorderStyle.Thin).ApllyStyle(ws.Cells[_Row, _Column]);
                        retorno = _Row;
                        firstRow = _Row;

                    }
                    else
                    {

                        ws.Cells[_Row, _Column - resultsQuantity, _Row, _Column-1].Copy(ws.Cells[_Row, _Column, _Row, _Column + (resultsQuantity-1)], OfficeOpenXml.ExcelRangeCopyOptionFlags.ExcludeFormulas);

                        object[,] arr = new object[1, resultsQuantity];
                        arr[0, 0] = orcado;
                        arr[0, 1] = forcast;
                        arr[0, 2] = realizado;
                        arr[0, 3] = orcadoVersusRealizado;
                        arr[0, 4] = forcastVersusRealizado;
                        arr[0, 5] = percentualOrcadoVersusRealizado;
                        arr[0, 6] = percentualForcastVersusRealizado;

                        ws.Cells[_Row, _Column, _Row, _Column + resultsQuantity].Value = arr;
                        _Column = _Column + (resultsQuantity-1);
                    }
                }
                else
                {
                    if (_Column == 2)
                    {
                        retorno = firstRow;
                        ws.Cells[firstRow , _Column, firstRow , 200].Copy(ws.Cells[_Row, _Column, _Row, _Column + 200], OfficeOpenXml.ExcelRangeCopyOptionFlags.ExcludeFormulas);
                    }

                    object[,] arr = new object[1, resultsQuantity];
                    arr[0, 0] = orcado;
                    arr[0, 1] = forcast;
                    arr[0, 2] = realizado;
                    arr[0, 3] = orcadoVersusRealizado;
                    arr[0, 4] = forcastVersusRealizado;
                    arr[0, 5] = percentualOrcadoVersusRealizado;
                    arr[0, 6] = percentualForcastVersusRealizado;

                    ws.Cells[_Row, _Column, _Row, _Column + resultsQuantity].Value = arr;
                    _Column = _Column + (resultsQuantity - 1);
                }

            });

            var subLevels = line.filhos;
            //var subLevels = line.data.Children;

            if (subLevels != null)
            {
                var pointerRow = 1;
                var ret = 0;
                foreach (var subLevel in subLevels)
                {
                    // subLevel.data.Children = line.data.Children.Where(x => x.data.IdPai.Contains(string.Format("{0}_{1}", subLevel.data.IdPai, subLevel.data.Cod))).ToList();
                    ret = CreateLineQuick(subLevel, type, periods, pointerRow++,ret);
                }
            }
            return retorno;
        }




        private string GetBackgroundColor(int level)
        {
            string color = "";

            if (level < 1)
                return "#FFF";

            int bgColor = 255 - (level * this.fator);
            bgColor = bgColor < 0 ? 0 : bgColor;
        Color myColor = Color.FromArgb(bgColor, bgColor, (bgColor + 10));
            color = "#" + myColor.R.ToString("X2") + myColor.G.ToString("X2") + myColor.B.ToString("X2");

            return color;
        }


    }
}
