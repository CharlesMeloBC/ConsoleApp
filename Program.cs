using OfficeOpenXml;

class Program
{
    static void Main(string[] args)
    {
        // Defina o contexto da licença
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Ou LicenseContext.Commercial se aplicável

        // Caminho do arquivo .xlsx
        var filePath = @"C:\Users\Charles\Desktop\Matriculas\DB.xlsx";

        // Carregar o arquivo
        FileInfo fileInfo = new FileInfo(filePath);
        using (var package = new ExcelPackage(fileInfo))
        {
            // Selecionar a primeira planilha
            var worksheet = package.Workbook.Worksheets[0];

            //// Ler dados de uma célula específica
            //var cellValue = worksheet.Cells[2,2].Text; // Linha 1, Coluna 1
            //Console.WriteLine(cellValue);

            // Para ler um intervalo de células
            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    Console.Write(worksheet.Cells[row, col].Text + "\t");
                }
                Console.WriteLine();
            }
        }
    }
}
