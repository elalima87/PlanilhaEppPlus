using System;
using System.Web.UI;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Data;
using System.Text;

using System.Web.Services;



namespace PlanilhaEppPlus
{
  public partial class _Default : Page
  {
    string TipoOperacao = "";
    string Relatorio = "";

    protected void Page_Load(object sender, EventArgs e)
    {
      if (Request["pTipoOperacao"] != "")
      {
        if (Request["pTipoOperacao"] != null)
        {
          TipoOperacao = Request["pTipoOperacao"].ToString();
        }
      }

      if (Request["pRelatorio"] != "")
      {
        if (Request["pRelatorio"] != null)
        {
          Relatorio = Request["pRelatorio"].ToString();
        }
      }

      if ((TipoOperacao == "X") && (Relatorio == "Rel1"))
      {
        GerarPlanilha();
      }

      if ((TipoOperacao == "X") && (Relatorio == "Rel2"))
      {
        GerarPlanilhaRel2();
      }

    }

    public static String columnName(long columnNumber)
    {
      StringBuilder retVal = new StringBuilder();
      int x = 0;
      for (int n = (int)(Math.Log(25 * (columnNumber + 1)) / Math.Log(26)) - 1; n >= 0; n--)
      {
        x = (int)((Math.Pow(26, (n + 1)) - 1) / 25 - 1);
        if (columnNumber > x)
          retVal.Append(System.Convert.ToChar((int)(((columnNumber - x - 1) / Math.Pow(26, n)) % 26 + 65)));
      }
      return retVal.ToString();
    }

    private ExcelPackage GerarPlanilha()
    {
      try
      {

        DataTable dTable = new DataTable();
        dTable = Relatorio1();

        //Remove coluna do dataTable
        dTable.Columns.Remove("ID_ATIVIDADE");
        //Renomeia coluna do dataTable
        dTable.Columns["TOTAL"].ColumnName = "Total de notas";

        // criando o arquivo:
        // criando uma planilha neste arquivo e obtendo a referência para meu código operá-la. ou arquivoExcel.Workbook.Worksheets[index];
        ExcelPackage wkBook = new ExcelPackage();
        //Create the worksheet
        ExcelWorksheet planilha = wkBook.Workbook.Worksheets.Add("Teste");

        System.Drawing.Color CorTitulo = System.Drawing.Color.FromArgb(0, 86, 150);
        System.Drawing.Color CorTotal = System.Drawing.Color.FromArgb(119, 130, 130);
        System.Drawing.Color CorFiltro = System.Drawing.Color.FromArgb(189, 215, 238);

        string ColumnFim = columnName(dTable.Columns.Count);

        planilha.Cells["A1"].LoadFromDataTable(dTable, true);

        //Obter a contagem de linhas e colunas
        var start = planilha.Dimension.Start;
        var end = planilha.Dimension.End;

        //Cabeçalho
        planilha.Cells["A1:" + ColumnFim + "1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        planilha.Cells["A1:" + ColumnFim + "1"].Style.Fill.BackgroundColor.SetColor(CorTitulo);
        planilha.Cells["A1:" + ColumnFim + "1"].Style.Font.Color.SetColor(System.Drawing.Color.White);
        planilha.Cells["A:" + end.Row.ToString()].AutoFitColumns();
        //  Adicionar filtros a linhas
        planilha.Cells["A1:" + ColumnFim + "1"].AutoFilter = true;

        //Embrulhe o texto
        planilha.Cells["A1:" + ColumnFim + "1"].Style.WrapText = true;

        planilha.Cells.Style.Font.Name = "Arial";      // Fonte Arial no documento inteiro
        planilha.Cells.Style.Font.Size = 11;           // Aplicando tamanho 11 no documento inteiro


        //      Atribuir cor de plano de fundo às células
        //rowRngprogramParamsRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
        //      rowRngprogramParamsRange.Style.Fill.BackgroundColor.SetColor(Color.DarkRed);
        //      Definir cor da fonte
        //rowRngprogramParamsRange.Style.Font.Color.SetColor(Color.Red);
        //      Definir alinhamento horizontal e vertical
        //columnHeaderRowRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //      columnHeaderRowRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;




        //Crie e configure uma única célula
        //using (var shortNameCell = locationWorksheet.Cells[rowToPop, SHORTNAME_BYDCBYLOC_COL])
        //{
        //  shortNameCell.Value = "Short Name";
        //  shortNameCell.Style.WrapText = false;
        //  shortNameCell.Style.Font.Size = 12;
        //}

        //      Adicionar uma imagem a uma folha
        //Para adicionar uma imagem a uma célula específica, faça o seguinte:

        //Ocultar código de cópia
        //string imgPath = @"C:\misc\yourImage.png"; AddImage(locationWorksheet, 1, 5, imgPath);
        //      AddImage(locationWorksheet, 1, 5, imgPath);
        //. . .
        //private void AddImage(ExcelWorksheet oSheet, int rowIndex, int colIndex, string imagePath)
        //    {
        //      Bitmap image = new Bitmap(imagePath);
        //      {
        //        var excelImage = oSheet.Drawings.AddPicture("Platypus Logo", image);
        //        excelImage.From.Column = colIndex - 1;
        //        excelImage.From.Row = rowIndex - 1;
        //        excelImage.SetSize(108, 84);
        //        excelImage.From.ColumnOff = Pixel2MTU(2);
        //        excelImage.From.RowOff = Pixel2MTU(2);
        //      }
        //    }

        //    public int Pixel2MTU(int pixels)
        //    {
        //      int mtus = pixels * 9525;
        //      return mtus;
        //    }


        //      Auto ou ajustar manualmente colunas
        //// autofit all columns
        //deliveryPerformanceWorksheet.Cells[deliveryPerformanceWorksheet.Dimension.Address].AutoFitColumns();
        //      customerWorksheet.Cells.AutoFitColumns();

        //      // autofit a specified range of columns
        //      locationWorksheet.Cells["A:C"].AutoFitColumns();

        //      // manually assign widths of specififed columns
        //      locationWorksheet.Column(4).Width = 3.14;
        //      locationWorksheet.Column(5).Width = 14.33;
        //      locationWorksheet.Column(6).Width = 186000000.00;
        //      Definir altura da linha
        //deliveryPerformanceWorksheet.Row(curDelPerfRow + 1).Height = HEIGHT_FOR_DELIVERY_PERFORMANCE_TOTALS_ROW;


        //      Adicionar bordas a um intervalo
        //using (var entireSheetRange = locationWorksheet.Cells[6, 1, locationWorksheet.Dimension.End.Row, 6])
        //      {
        //        entireSheetRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        //        entireSheetRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
        //        entireSheetRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
        //        entireSheetRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
        //        entireSheetRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        //      }

        // Escrevendo tabela
        //foreach (Livro livro in livros)
        //{
        //  coluna = 1;
        //  planilha.Cells[linha, coluna++].Value = livro.Categoria.Descricao;
        //  planilha.Cells[linha, coluna++].Value = livro.Titulo;
        //  planilha.Cells[linha, coluna++].Value = livro.GetNomesAutores();

        //  planilha.Cells[linha, coluna].Style.Numberformat.Format = "R$ ###.00"; // formatando para dinheiro, ANTES de escrever na célula
        //  planilha.Cells[linha, coluna++].Value = livro.Preco;

        //  linha++;
        //}

        // Auto ajustando a largura das células da tabela
        // planilha.Cells[linhaInicioTabela, coluna, linha - 1, coluna + 3].AutoFitColumns();


        //      Definir um formato numérico para uma célula
        //Ocultar   código de cópia
        //monOrdersCell.Style.Numberformat.Format = "0"; // some other possibilities are "#,##0.00";  
        //                                               // "#,##0"; "\"$\"#,##0.00;[Red]\"$\"#,##0.00"; 
        //                                               // "_($* #,##0.00_);_($* (#,##0.00);_($* \"\" - \"\"??_);
        //                                               // _(@_)";  "0.00";
        //      Adicione uma fórmula a uma célula
        //Ocultar código de cópia
        //using (var totalTotalOrdersCell = deliveryPerformanceWorksheet.Cells[curDelPerfRow + 1,
        //       TOTAL_ORDERS_COLUMN])
        //      {
        //        totalTotalOrdersCell.Style.Numberformat.Format = "#,##0";
        //        totalTotalOrdersCell.Formula = string.Format("SUM(J{0}:J{1})", FIRST_DATA_ROW, curDelPerfRow - 1);
        //        totalTotalOrdersCell.Calculate();
        //        // Note that EPPlus apparently differs from Excel Interop in that there is no "="
        //        // at the beginning of the formula, e.g. it does not start "=SUM("
        //        // Another way (rather than using a defined range, as above) is: 
        //        // deliveryPerformanceWorksheet.Cells["C4"].Formula = "SUM(C2:C3)";
        //      }

        //      Soma manual de um intervalo de linhas em uma coluna
        //Eu tive uma ocasião em que a Fórmula não funcionou para mim e tive que "forçar a força"; aqui está como eu fiz assim:

        //Ocultar código de cópia
        //totalOccurrencesCell.Value = SumCellVals(SUMMARY_TOTAL_OCCURRENCES_COL, FIRST_SUMMARY_DATA_ROW, rowToPopulate - 1);
        //. . .
        //private string SumCellVals(int colNum, int firstRow, int lastRow)
        //    {
        //      double runningTotal = 0.0;
        //      double currentVal;
        //      for (int i = firstRow; i <= lastRow; i++)
        //      {
        //        using (var sumCell = priceComplianceWorksheet.Cells[i, colNum])
        //        {
        //          currentVal = Convert.ToDouble(sumCell.Value);
        //          runningTotal = runningTotal + currentVal;
        //        }
        //      }
        //      return runningTotal.ToString();
        //    }
        //    Para somar ints em vez de números reais, basta alterá-lo para usar ints em vez de duplos.


        //      Esconder uma Fila
        //Ocultar código de cópia
        //yourWorksheet.Row(_lastRowAdded).Hidden = true;
        //    Ocultar linhas de grade em uma folha
        //    Ocultar   código de cópia
        //    priceComplianceWorksheet.View.ShowGridLines = false;
        //    Especifique uma linha de repetição para imprimir em páginas subseqüentes
        //Ocultar código de cópia
        //prodUsageWorksheet.PrinterSettings.RepeatRows = new ExcelAddress(String.Format("${0}:${0}", COLUMN_HEADING_ROW));




        // salvando e fechando o arquivo
        Response.Clear();
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        Response.AddHeader("content-disposition", "attachment;  filename=Teste.xlsx");
        Response.BinaryWrite(wkBook.GetAsByteArray());
        Response.Flush();
        Response.End();
      }

      catch (Exception ex)
      {

      }
      return null;


    }

    public DataTable Relatorio1()
    {
      DataTable dTable = new DataTable();

      dTable.Columns.Add("DS_QUESTIONARIO", typeof(string));
      dTable.Columns.Add("AVALIACAO", typeof(string));
      dTable.Columns.Add("AVALIADO", typeof(string));
      dTable.Columns.Add("DS_COMPETENCIA", typeof(string));
      dTable.Columns.Add("DS_GRUPO_ATIVIDADE", typeof(string));
      dTable.Columns.Add("DS_ATIVIDADE", typeof(string));
      dTable.Columns.Add("ID_ATIVIDADE", typeof(string));
      dTable.Columns.Add("NR_ORDEM", typeof(int));
      dTable.Columns.Add("TOTAL", typeof(string));


      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PR VISITA (PLANEJAMENTO)", "An lise", "01. 01 - Auditoria: Interpreta‡Æo da informa‡Æo dispon¡vel", "1241", "1", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PR VISITA (PLANEJAMENTO)", "An lise", "01. 02 - Matriz de decisÆo: Conquistar / Ampliar / Blindar / Expandir", "1242", "2", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PR VISITA (PLANEJAMENTO)", "An lise", "01. 03 - Perfil M‚dico: Identifica‡Æo de perfil e forma/adapta‡Æo a abordagem rep X cliente", "1243", "3", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PR VISITA (PLANEJAMENTO)", "Objetivo", "02. 01 - Produto Foco 1: Escolha do produto de maior abordagem", "1244", "4", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PR VISITA (PLANEJAMENTO)", "Objetivo", "02. 02 - Produto Foco 2: Segunda op‡Æo de abordagem", "1245", "5", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PR VISITA (PLANEJAMENTO)", "Plano de abordagem", "03. 01 - Como ser  a abertura: Defini‡Æo do modelo e tipo de questionamento", "1296", "6", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PR VISITA (PLANEJAMENTO)", "Plano de abordagem", "03. 02 - Recursos a utilizar: Tablet / Separatas / Amostras / Brindes", "1247", "7", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "VISITA", "Abertura", "04. 01 Temas: Consegue introduzir o assunto?", "1248", "8", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "VISITA", "Sondagens", "05. 01 - Sondagem Ativa: A sondagem foi adequada para atingir o objetivo proposto?", "1249", "9", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "VISITA", "Apresenta‡Æo de Produto", "06. 01 - Caracter¡stica: Informa para o que serve o produto?", "1251", "10", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "VISITA", "Apresenta‡Æo de Produto", "06. 02 - Benef¡cios: Informa os ganhos obtidos pelo m‚dico / paciente?", "1252", "11", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "VISITA", "Apresenta‡Æo de Produto", "06. 03 - Vantagens: Diferencia seu produto do concorrente?", "1253", "12", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "VISITA", "Apresenta‡Æo de Produto", "06. 04 - Provas: Apresenta estudos ou outros recursos que possam comprovar caracter¡sticas e benef¡cios?", "1254", "13", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "VISITA", "Obje‡äes", "07. 01 - Levantamento: Extrai informa‡äes contr rias ao uso de seu produto?", "1255", "14", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "VISITA", "Obje‡äes", "07. 02 - Manejo / Contorno: Atende com assertividade os questionamentos?", "1256", "15", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "VISITA", "Compromisso", "08. 01 - Acordo: Eleva o cliente na escala de decisÆo?", "1257", "16", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PàS VISITA (ANµLISE CRÖTICA) ", "Cumprimento dos Objetivos", "09. 01 - Plano X Realizado: Cumpriu fielmente o acordado na Pr‚ Visita?", "1258", "17", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PàS VISITA (ANµLISE CRÖTICA) ", "Cumprimento dos Objetivos", "09. 02 - Assertividade Aloca‡Æo Recursos - Entregou os materiais previstos pelo MKT ? A quantidade de Ags investida foi alinhada a Pr‚ Visita ", "1259", "18", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PàS VISITA (ANµLISE CRÖTICA) ", "Registro de Achados", "10. 01 - Anota‡äes: Coment rios qualitativos da visita", "1297", "19", "A");
      dTable.Rows.Add("PAAR 2017: MEDICO", "PAAR 2017: MEDICO", "ANTONIO JOSE DE OLIVEIRA", "PàS VISITA (ANµLISE CRÖTICA) ", "Registro de Achados", "10. 02 - Objetivo Pr¢xima Visita: Constr¢i objetivamente o OPV com base no registro de achados?", "1261", "20", "A");

      return dTable;
    }


    private void AddImage(ExcelWorksheet oSheet, int rowIndex, int colIndex, string imagePath)
    {
      Bitmap image = new Bitmap(imagePath);
      {
        var excelImage = oSheet.Drawings.AddPicture("Platypus Logo", image);
        excelImage.From.Column = colIndex - 5;
        excelImage.From.Row = rowIndex - 1;
        excelImage.SetSize(90, 50);
        excelImage.From.ColumnOff = Pixel2MTU(2);
        excelImage.From.RowOff = Pixel2MTU(2);
      }
    }

    public int Pixel2MTU(int pixels)
    {
      int mtus = pixels * 9525;
      return mtus;
    }



    private ExcelPackage GerarPlanilhaRel2()
    {
      try
      {

        DataTable dTable = new DataTable();
        dTable = Relatorio2();

        //Renomeia coluna do dataTable
        dTable.Columns["ID"].ColumnName = "Id";
        dTable.Columns["NOME"].ColumnName = "Nome";
        dTable.Columns["EMAIL"].ColumnName = "E-mail";
        dTable.Columns["DT_NASCIMENTO"].ColumnName = "Data nascimento";
        dTable.Columns["TOTAL_VENDAS"].ColumnName = "Total de vendas";

        // criando o arquivo:
        // criando uma planilha neste arquivo e obtendo a referência para meu código operá-la. ou arquivoExcel.Workbook.Worksheets[index];
        ExcelPackage wkBook = new ExcelPackage();
        //Create the worksheet
        ExcelWorksheet planilha = wkBook.Workbook.Worksheets.Add("Força de vendas");

        System.Drawing.Color CorTitulo = System.Drawing.Color.FromArgb(0, 86, 150);
        System.Drawing.Color CorTotal = System.Drawing.Color.FromArgb(119, 130, 130);
        System.Drawing.Color CorFiltro = System.Drawing.Color.FromArgb(189, 215, 238);

        string ColumnFim = columnName(dTable.Columns.Count);

        planilha.Cells["A5"].LoadFromDataTable(dTable, true);

        //Obter a contagem de linhas e colunas
        var start = planilha.Dimension.Start;
        var end = planilha.Dimension.End;

        //      Atribuir cor de plano de fundo às células
        //rowRngprogramParamsRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
       // planilha.Cells.Style.Fill.BackgroundColor.SetColor(Color.White);

        //Cabeçalho
        planilha.Cells["A5:" + ColumnFim + "5"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        planilha.Cells["A5:" + ColumnFim + "5"].Style.Fill.BackgroundColor.SetColor(CorTitulo);
        planilha.Cells["A5:" + ColumnFim + "5"].Style.Font.Color.SetColor(System.Drawing.Color.White);
        planilha.Cells["A:" + end.Row.ToString()].AutoFitColumns();
        //  Adicionar filtros a linhas
        planilha.Cells["A5:" + ColumnFim + "5"].AutoFilter = true;
        
        planilha.Cells.Style.Font.Name = "Arial";      // Fonte Arial no documento inteiro
        planilha.Cells.Style.Font.Size = 11;           // Aplicando tamanho 11 no documento inteiro
        
        //Adicionar uma imagem a uma folha
        //Para adicionar uma imagem a uma célula específica, faça o seguinte:
        string imgPath = @"C:\Users\rafaela.pinheiro\source\repos\PlanilhaEppPlus\PlanilhaEppPlus\img\Logo_teste.png";
        AddImage(planilha, 1, 5, imgPath);

        //      Adicionar bordas a um intervalo
        using (var entireSheetRange = planilha.Cells[6, 1,planilha.Dimension.End.Row, 5])
        {
          entireSheetRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
          entireSheetRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
          entireSheetRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
          entireSheetRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
          entireSheetRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        }

        

        //Crie e configure uma única célula
        //using (var shortNameCell = locationWorksheet.Cells[rowToPop, SHORTNAME_BYDCBYLOC_COL])
        //{
        //  shortNameCell.Value = "Short Name";
        //  shortNameCell.Style.WrapText = false;
        //  shortNameCell.Style.Font.Size = 12;
        //}




        // Escrevendo tabela
        //foreach (Livro livro in livros)
        //{
        //  coluna = 1;
        //  planilha.Cells[linha, coluna++].Value = livro.Categoria.Descricao;
        //  planilha.Cells[linha, coluna++].Value = livro.Titulo;
        //  planilha.Cells[linha, coluna++].Value = livro.GetNomesAutores();

        //  planilha.Cells[linha, coluna].Style.Numberformat.Format = "R$ ###.00"; // formatando para dinheiro, ANTES de escrever na célula
        //  planilha.Cells[linha, coluna++].Value = livro.Preco;

        //  linha++;
        //}

        

        //      Definir um formato numérico para uma célula
        //Ocultar   código de cópia
        //monOrdersCell.Style.Numberformat.Format = "0"; // some other possibilities are "#,##0.00";  
        //                                               // "#,##0"; "\"$\"#,##0.00;[Red]\"$\"#,##0.00"; 
        //                                               // "_($* #,##0.00_);_($* (#,##0.00);_($* \"\" - \"\"??_);
        //                                               // _(@_)";  "0.00";
        //      Adicione uma fórmula a uma célula
        //Ocultar código de cópia
        //using (var totalTotalOrdersCell = deliveryPerformanceWorksheet.Cells[curDelPerfRow + 1,
        //       TOTAL_ORDERS_COLUMN])
        //      {
        //        totalTotalOrdersCell.Style.Numberformat.Format = "#,##0";
        //        totalTotalOrdersCell.Formula = string.Format("SUM(J{0}:J{1})", FIRST_DATA_ROW, curDelPerfRow - 1);
        //        totalTotalOrdersCell.Calculate();
        //        // Note that EPPlus apparently differs from Excel Interop in that there is no "="
        //        // at the beginning of the formula, e.g. it does not start "=SUM("
        //        // Another way (rather than using a defined range, as above) is: 
        //        // deliveryPerformanceWorksheet.Cells["C4"].Formula = "SUM(C2:C3)";
        //      }

        //      Soma manual de um intervalo de linhas em uma coluna
        //Eu tive uma ocasião em que a Fórmula não funcionou para mim e tive que "forçar a força"; aqui está como eu fiz assim:

        //Ocultar código de cópia
        //totalOccurrencesCell.Value = SumCellVals(SUMMARY_TOTAL_OCCURRENCES_COL, FIRST_SUMMARY_DATA_ROW, rowToPopulate - 1);
        //. . .
        //private string SumCellVals(int colNum, int firstRow, int lastRow)
        //    {
        //      double runningTotal = 0.0;
        //      double currentVal;
        //      for (int i = firstRow; i <= lastRow; i++)
        //      {
        //        using (var sumCell = priceComplianceWorksheet.Cells[i, colNum])
        //        {
        //          currentVal = Convert.ToDouble(sumCell.Value);
        //          runningTotal = runningTotal + currentVal;
        //        }
        //      }
        //      return runningTotal.ToString();
        //    }
        //    Para somar ints em vez de números reais, basta alterá-lo para usar ints em vez de duplos.


        //      Esconder uma Fila
        //Ocultar código de cópia
        //yourWorksheet.Row(_lastRowAdded).Hidden = true;
        //    Ocultar linhas de grade em uma folha
        //    Ocultar   código de cópia
        //    priceComplianceWorksheet.View.ShowGridLines = false;
        //    Especifique uma linha de repetição para imprimir em páginas subseqüentes
        //Ocultar código de cópia
        //prodUsageWorksheet.PrinterSettings.RepeatRows = new ExcelAddress(String.Format("${0}:${0}", COLUMN_HEADING_ROW));




        // salvando e fechando o arquivo
        Response.Clear();
        Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        Response.AddHeader("content-disposition", "attachment;  filename=Força_de_vendas.xlsx");
        Response.BinaryWrite(wkBook.GetAsByteArray());
        Response.Flush();
        Response.End();
      }

      catch (Exception ex)
      {

      }
      return null;


    }


    public DataTable Relatorio2()
    {
      DataTable dTable = new DataTable();


      dTable.Columns.Add("ID", typeof(int));
      dTable.Columns.Add("NOME", typeof(string));
      dTable.Columns.Add("EMAIL", typeof(string));
      dTable.Columns.Add("DT_NASCIMENTO", typeof(string));
      dTable.Columns.Add("TOTAL_VENDAS", typeof(string));

      dTable.Rows.Add("1","Eaton Berg","tellus@dui.co.uk","06/06/2019","R$ 100,00");
      dTable.Rows.Add("2","Rigel Palmer","Quisque.imperdiet@id.ca","20/12/2018","R$ 150,00");
      dTable.Rows.Add("3","Malachi Howe","enim.commodo@lacus.edu","29/07/2017","R$ 25,00");
      dTable.Rows.Add("4","Wallace Rollins","mollis.Phasellus.libero@consequatdolor.net","25/08/2018","R$ 154,00");
      dTable.Rows.Add("5","Yoshio Flynn","nec@loremeget.ca","17/08/2017","R$ 548,00");
      dTable.Rows.Add("6","Adam Mendez","Maecenas.libero@Quisquefringilla.ca","24/02/2018","R$ 45,00");
      dTable.Rows.Add("7","Tyler Gibbs","habitant@Sedet.com","01/07/2018","R$ 125,00");
      dTable.Rows.Add("8","Zeus Olsen","et.rutrum@Aliquamauctorvelit.net","29/07/2017","R$ 0,00");
      dTable.Rows.Add("9","Linus Goodman","lorem.Donec.elementum@id.edu","04/08/2018","R$ 875,00");
      dTable.Rows.Add("10","Bevis Wong","Ut.nec.urna@loremDonecelementum.com","31/10/2018","R$ 587,00");
      dTable.Rows.Add("11","Kareem Hancock","consectetuer.adipiscing.elit@metusfacilisislorem.net","06/10/2017","R$ 5,00");
      dTable.Rows.Add("12","Buckminster Roberson","risus.Donec.egestas@nisimagnased.net","10/05/2018","R$ 548,00");
      dTable.Rows.Add("13","Kyle Berg","mus@Vestibulumaccumsanneque.org","18/12/2018","R$ 57,00");
      dTable.Rows.Add("14","Jakeem Barrera","arcu.ac.orci@dolor.net","25/12/2017","R$ 4.587,00");
      dTable.Rows.Add("15","Silas Underwood","malesuada@tinciduntnequevitae.co.uk","31/12/2017","R$ 854,00");
      dTable.Rows.Add("16","Beau Peters","ligula.Aenean.gravida@ut.org","30/12/2017","R$ 54,00");
      dTable.Rows.Add("17","Rigel Barnes","eu@lacusvestibulumlorem.ca","04/11/2017","R$ 547,00");
      dTable.Rows.Add("18","Dieter Hernandez","enim.Etiam.imperdiet@fringilla.org","23/06/2019","R$ 5.478,00");
      dTable.Rows.Add("19","Dalton Hays","velit@euismodmauriseu.ca","02/05/2019","R$ 5,00");
      dTable.Rows.Add("20","Camden Mercado","gravida.mauris@diam.com","22/12/2018","R$ 589,00");
      dTable.Rows.Add("21","Timothy Reilly","sed.consequat@ipsumCurabiturconsequat.edu","15/02/2019","R$ 0,00");
      dTable.Rows.Add("22","Cain Barber","Praesent.eu.nulla@blandit.net","04/04/2019","R$ 0,00");
      dTable.Rows.Add("23","Dalton Carver","Phasellus.ornare@orcisemeget.co.uk","02/02/2018","R$ 477,00");
      dTable.Rows.Add("24","Aladdin Black","ipsum.dolor.sit@hendrerit.co.uk","24/10/2018","R$ 785,00");
      dTable.Rows.Add("25","Leo Barnett","tincidunt.nunc@mauris.co.uk","14/04/2019","R$ 887,00");
      dTable.Rows.Add("26","Prescott Blair","Duis.gravida.Praesent@rhoncusNullam.ca","31/05/2018","R$ 0,00");
      dTable.Rows.Add("27","Thaddeus Gentry","mollis.nec.cursus@id.org","09/11/2017","R$ 85,00");
      dTable.Rows.Add("28","August Collins","molestie.orci.tincidunt@odio.net","21/05/2018","R$ 57,00");
      dTable.Rows.Add("29","Grant Sutton","fames.ac@vitae.ca","06/11/2018","R$ 0,00");
      dTable.Rows.Add("30","Thane Chambers","velit.eu.sem@quis.ca","21/06/2018","R$ 0,00");
      dTable.Rows.Add("31","Jelani Lott","ultrices.posuere.cubilia@eleifendnecmalesuada.co.uk","16/09/2017","R$ 0,00");
      dTable.Rows.Add("32","Howard Hall","mi.eleifend@sollicitudinadipiscing.com","22/03/2018","R$ 0,00");
      dTable.Rows.Add("33","Nigel Bullock","bibendum.fermentum.metus@Cum.ca","01/10/2018","R$ 0,00");
      dTable.Rows.Add("34","Moses Mccullough","In.condimentum@dictumaugue.co.uk","21/01/2019","R$ 0,00");
      dTable.Rows.Add("35","Aidan Irwin","Nunc.sollicitudin@Aliquamerat.ca","09/06/2018","R$ 0,00");
      dTable.Rows.Add("36","Otto Webster","ac@MorbivehiculaPellentesque.org","13/03/2018","R$ 0,00");
      dTable.Rows.Add("37","Jonah Gross","luctus.felis@nequeetnunc.net","09/08/2018","R$ 0,00");
      dTable.Rows.Add("38","Ishmael Hensley","arcu@facilisis.co.uk","08/06/2018","R$ 0,00");
      dTable.Rows.Add("39","Armand Wood","Donec.dignissim@auctorveliteget.edu","25/06/2019","R$ 0,00");
      dTable.Rows.Add("40","Dustin Pratt","ac.mi.eleifend@Mauris.com","22/10/2017","R$ 0,00");
      dTable.Rows.Add("41","Igor Cervantes","urna.Vivamus@Namligula.com","27/06/2018","R$ 0,00");
      dTable.Rows.Add("42","Jack Hendrix","aliquam@risusDonecnibh.com","15/05/2019","R$ 87,00");
      dTable.Rows.Add("43","Wayne Vance","tincidunt.tempus.risus@eget.net","14/05/2019","R$ 47,00");
      dTable.Rows.Add("44","Duncan Carr","erat.Etiam@vestibulumloremsit.net","08/09/2017","R$ 875,00");
      dTable.Rows.Add("45","Magee Battle","massa@at.net","21/01/2018","R$ 54,00");
      dTable.Rows.Add("46","Merritt Faulkner","venenatis.a@ametnullaDonec.net","16/07/2018","R$ 21,00");
      dTable.Rows.Add("47","Kibo Dalton","Donec.felis.orci@Aliquamfringillacursus.net","12/05/2018","R$ 8,00");
      dTable.Rows.Add("48","Gavin Bond","tincidunt@mollisneccursus.com","25/09/2017","R$ 4,00");
      dTable.Rows.Add("49","Emmanuel Neal","bibendum.sed@eueuismod.net","22/05/2019","R$ 4,00");
      dTable.Rows.Add("50","Leroy Frank","Aliquam@blanditmattisCras.co.uk","05/05/2019","R$ 5,00");
      dTable.Rows.Add("51","Leo Ford","arcu.Vivamus.sit@tristiqueneque.net","12/02/2019","R$ 24,00");
      dTable.Rows.Add("52","Otto Dunlap","tristique@lectus.ca","25/04/2019","R$ 2,00");
      dTable.Rows.Add("53","Samson Mclaughlin","sociis.natoque.penatibus@pretium.org","17/01/2019","R$ 54,00");
      dTable.Rows.Add("54","Chester Mosley","felis@pedeacurna.edu","24/03/2018","R$ 87,00");
      dTable.Rows.Add("55","Orson Foreman","elit@elitelitfermentum.ca","16/12/2018","R$ 778,00");
      dTable.Rows.Add("56","Wesley Schroeder","aliquet.lobortis.nisi@idante.edu","27/11/2018","R$ 87,00");
      dTable.Rows.Add("57","Igor Wilder","dolor.vitae@purusmaurisa.org","14/05/2019","R$ 0,00");
      dTable.Rows.Add("58","Carter Wolf","sodales.purus.in@molestie.edu","26/02/2018","R$ 0,00");
      dTable.Rows.Add("59","Castor Harvey","et.euismod@nisinibh.ca","18/08/2017","R$ 75,00");
      dTable.Rows.Add("60","Prescott Sellers","per.conubia.nostra@velvulputateeu.net","13/12/2017","R$ 4,00");
      dTable.Rows.Add("61","Magee James","Nullam@magna.ca","03/01/2018","R$ 5,00");
      dTable.Rows.Add("62","Carson Holman","malesuada.fames.ac@loremauctor.com","22/04/2019","R$ 4,00");
      dTable.Rows.Add("63","Slade Allison","mi.lacinia.mattis@atarcu.com","05/09/2017","R$ 1,00");
      dTable.Rows.Add("64","Garth Rojas","ipsum.ac.mi@ametanteVivamus.co.uk","17/04/2018","R$ 1,00");
      dTable.Rows.Add("65","Ryder Salas","Mauris.vestibulum.neque@Etiamlaoreet.com","04/12/2018","R$ 478,00");
      dTable.Rows.Add("66","Ronan Mcguire","sagittis.semper.Nam@amagnaLorem.edu","13/06/2019","R$ 4,00");
      dTable.Rows.Add("67","Geoffrey Gutierrez","molestie.pharetra@dictumeu.net","28/07/2017","R$ 47,00");
      dTable.Rows.Add("68","Kamal West","nec.malesuada.ut@temporestac.net","27/09/2018","R$ 47,00");
      dTable.Rows.Add("69","Channing Bean","diam.vel.arcu@Donecegestas.edu","06/08/2018","R$ 3,00");
      dTable.Rows.Add("70","Jesse Small","Nam.tempor@vel.co.uk","16/12/2018","R$ 658,00");
      dTable.Rows.Add("71","Dolan Wells","lobortis.nisi@arcuSed.ca","19/01/2019","R$ 87,00");
      dTable.Rows.Add("72","Reese Turner","fringilla@Etiamgravida.edu","07/12/2018","R$ 7,00");
      dTable.Rows.Add("73","Hashim Guerra","metus.eu@consequatpurusMaecenas.org","14/12/2018","R$ 547,00");
      dTable.Rows.Add("74","Isaac Rodgers","nibh.Quisque.nonummy@orci.ca","14/01/2018","R$ 47,00");
      dTable.Rows.Add("75","Elvis Sellers","neque.sed@in.com","18/04/2019","R$ 4,00");
      dTable.Rows.Add("76","Duncan Gallegos","Pellentesque.habitant.morbi@gravida.com","31/05/2018","R$ 0,00");
      dTable.Rows.Add("77","Carlos Barnes","mollis.Phasellus.libero@sit.org","10/11/2018","R$ 0,00");
      dTable.Rows.Add("78","Ivan Gardner","Phasellus.dapibus.quam@etultrices.co.uk","25/05/2018","R$ 0,00");
      dTable.Rows.Add("79","Zahir Hayes","ante@lobortistellusjusto.net","26/08/2018","R$ 545,00");
      dTable.Rows.Add("80","Robert Stanley","Donec.est@Duis.org","25/02/2019","R$ 214,00");
      dTable.Rows.Add("81","Chadwick Price","adipiscing@neque.edu","19/04/2019","R$ 6.578,00");
      dTable.Rows.Add("82","Vincent Barton","In@egestas.edu","24/05/2018","R$ 47,00");
      dTable.Rows.Add("83","Vernon Carlson","Cum.sociis@nostraper.org","05/05/2019","R$ 5,00");
      dTable.Rows.Add("84","Keaton Simmons","viverra@tinciduntcongueturpis.edu","14/09/2017","R$ 85,00");
      dTable.Rows.Add("85","Alden Gould","imperdiet.ullamcorper@In.ca","30/09/2018","R$ 74,00");
      dTable.Rows.Add("86","Conan Drake","amet.massa.Quisque@ornare.ca","23/02/2018","R$ 596,00");
      dTable.Rows.Add("87","Ivan Roth","dictum.Proin.eget@mauriserateget.org","18/02/2019","R$ 577,00");
      dTable.Rows.Add("88","Stone Wiley","amet@quistristique.ca","19/05/2018","R$ 287,00");
      dTable.Rows.Add("89","Oliver Christian","Phasellus.nulla.Integer@Phasellusvitaemauris.edu","27/05/2019","R$ 247,00");
      dTable.Rows.Add("90","Alexander Holden","viverra@gravidaAliquamtincidunt.net","31/08/2017","R$ 212,00");
      dTable.Rows.Add("91","Dalton Gay","metus@Donec.org","26/04/2018","R$ 214,00");
      dTable.Rows.Add("92","Len Estes","lacinia.Sed.congue@ipsum.ca","14/07/2019","R$ 4,00");
      dTable.Rows.Add("93","Timothy Crane","tincidunt.neque@Nuncmauris.edu","31/08/2018","R$ 0,00");
      dTable.Rows.Add("94","Nathan Frank","cursus.in@nislarcu.edu","03/03/2018","R$ 0,00");
      dTable.Rows.Add("95","Timon Wilson","consequat.enim@montes.ca","08/05/2019","R$ 547,00");
      dTable.Rows.Add("96","Sebastian Battle","in.hendrerit.consectetuer@enimSuspendisse.com","30/12/2017","R$ 547,00");
      dTable.Rows.Add("97","Hoyt Hutchinson","faucibus@tempor.com","24/08/2017","R$ 234,00");
      dTable.Rows.Add("98","Colton Neal","ipsum@eu.co.uk","23/11/2018","R$ 578,00");
      dTable.Rows.Add("99","Solomon Warner","nisi.a@maurisipsum.com","17/07/2019","R$ 878,00");
      dTable.Rows.Add("100","Tad Barnes","vestibulum.neque@imperdietdictum.org","27/06/2018","R$ 578,00");
      
      return dTable;
    }




  }
}