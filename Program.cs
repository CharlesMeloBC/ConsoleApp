using OfficeOpenXml;

class Program
{
    public class Matricula
    {
        public int Numero { get; set; } = 0;
        public string Solicitante { get; set; } = "";
        public string Status { get; set; } = "";
        public string Responsavel { get; set; } = "";
    }

    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        var filePath = @"C:\Users\Charles\Desktop\Matriculas\DB.xlsx";

        List<Matricula> matriculas = new List<Matricula>();

        FileInfo fileInfo = new FileInfo(filePath);
        using (var package = new ExcelPackage(fileInfo))
        {
            var worksheet = package.Workbook.Worksheets[0];

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var matricula = new Matricula
                {
                    Numero = int.TryParse(worksheet.Cells[row, 1].Text, out int numero) ? numero : 0, // Numero
                    Solicitante = worksheet.Cells[row, 2].Text, // Solicitante
                    Status = worksheet.Cells[row, 3].Text, // Status
                    Responsavel = worksheet.Cells[row, 4].Text // Responsável 
                };

                matriculas.Add(matricula);
            }
        }

        var arrayDeMatricula = matriculas.ToArray();

        Console.WriteLine("Número\tSolicitante\tStatus\tResponsável");
        Console.WriteLine("----------------------------------------------------");
        foreach (var matricula in arrayDeMatricula)
        {
            Console.WriteLine($"{matricula.Numero}\t{matricula.Solicitante}\t{matricula.Status}\t{matricula.Responsavel}");
        }
    }
}
