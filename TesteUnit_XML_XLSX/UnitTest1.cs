using ClosedXML.Excel;
namespace TesteUnit_XML_XLSX
{
    public class UnitTest1
    {
        [Fact]
        public async Task ExcelToXml_DeveRetornarMensagem_QuandoNomeEhNulo()
        {
            var resultado = await Program.ExcelToXml(null, "qualquercoisa");
            Assert.Equal("Arquivo Excel não encontrado.", resultado);
        }

        [Fact]
        public async Task ExcelToXml_DeveRetornarMensagem_QuandoNomeContemXlsx()
        {
            var resultado = await Program.ExcelToXml("arquivo.xlsx", "qualquercoisa");
            Assert.Equal("Arquivo Excel não encontrado.", resultado);
        }

        [Fact]
        public async Task ExcelToXml_DeveRetornarMensagem_QuandoBase64EhNulo()
        {
            var resultado = await Program.ExcelToXml("arquivo", null);
            Assert.Equal("Arquivo Excel não encontrado.", resultado);
        }

        [Fact]
        public async Task ExcelToXml_DeveConverterExcelParaXmlCorretamente()
        {
            // Arrange: cria um arquivo Excel em memória
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Planilha1");
            worksheet.Cell(1, 1).Value = "Nome";
            worksheet.Cell(1, 2).Value = "Idade";
            worksheet.Cell(2, 1).Value = "João";
            worksheet.Cell(2, 2).Value = "30";
            worksheet.Cell(3, 1).Value = "Maria";
            worksheet.Cell(3, 2).Value = "25";

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            var base64 = Convert.ToBase64String(ms.ToArray());

            // Act
            var resultadoBase64 = await Program.ExcelToXml("arquivo", base64);

            // Assert: decodifica o XML e verifica o conteúdo
            var xmlBytes = Convert.FromBase64String(resultadoBase64);
            var xmlString = System.Text.Encoding.UTF8.GetString(xmlBytes);

            Assert.Contains("<Registros>", xmlString);
            Assert.Contains("<Nome>João</Nome>", xmlString);
            Assert.Contains("<Idade>30</Idade>", xmlString);
            Assert.Contains("<Nome>Maria</Nome>", xmlString);
            Assert.Contains("<Idade>25</Idade>", xmlString);
        }
    }
}
