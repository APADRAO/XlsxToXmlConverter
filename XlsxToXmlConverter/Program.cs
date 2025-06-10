using ClosedXML.Excel;
using System.Xml.Linq;
/// <summary>
/// Classe para conversão de arquivos XML para Excel (XLSX) e vice-versa.
/// </summary>

public class Program
{
    /// <summary>
    /// Converte um arquivo XML para Excel (XLSX) de forma assíncrona.
    /// </summary>
    /// <param name="nome"></param>
    /// <param name="base64"></param>
    /// <returns></returns>
    public static async Task<string> XmlToExcelAsinc(string nome, string base64)
    {
        if (string.IsNullOrEmpty(nome) || nome.ToLower().Contains(".xml") || string.IsNullOrEmpty(base64))
        {
            return "Arquivo XML não encontrado.";
        }

        var doc = await XDocument.LoadAsync(new MemoryStream(Convert.FromBase64String(base64)), System.Xml.Linq.LoadOptions.None, CancellationToken.None);
        var root = doc.Root;
        if (root == null)
        {
            return "XML inválido.";
        }

        var rows = new List<Dictionary<string, string>>();
        FlattenElement(root, new Dictionary<string, string>(), rows, "");

        var headers = rows.SelectMany(r => r.Keys).Distinct().ToList();

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Dados");

        for (int i = 0; i < headers.Count; i++)
            worksheet.Cell(1, i + 1).Value = headers[i];

        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            for (int j = 0; j < headers.Count; j++)
            {
                row.TryGetValue(headers[j], out var value);
                worksheet.Cell(i + 2, j + 1).Value = value ?? "";
            }
        }
        var memory = new MemoryStream();
        // Reset the position to the beginning of the stream before converting to Base64
        memory.Position = 0;

        // Save the workbook to the memory stream
        workbook.SaveAs(memory);

        // Convert the memory stream to a Base64 string
        await Task.CompletedTask;
        return Convert.ToBase64String(memory.ToArray());
    }

    /// <summary>
    /// Flatten an XElement into a list of dictionaries representing rows for Excel.
    /// </summary>
    /// <param name="element"></param>
    /// <param name="currentRow"></param>
    /// <param name="rows"></param>
    /// <param name="prefix"></param>
    protected static void FlattenElement(XElement element, Dictionary<string, string> currentRow, List<Dictionary<string, string>> rows, string prefix)
    {
        foreach (var attr in element.Attributes())
        {
            currentRow[$"{prefix}@{attr.Name.LocalName}"] = attr.Value;
        }

        var children = element.Elements().ToList();
        if (!children.Any())
        {
            currentRow[prefix.TrimEnd('/')] = element.Value;
        }
        else
        {
            bool isRepeated = children.GroupBy(c => c.Name.LocalName).Any(g => g.Count() > 1);
            foreach (var group in children.GroupBy(c => c.Name.LocalName))
            {
                var elements = group.ToList();
                for (int i = 0; i < elements.Count; i++)
                {
                    var child = elements[i];
                    var newRow = new Dictionary<string, string>(currentRow);
                    var childPrefix = isRepeated ? $"{prefix}{child.Name.LocalName}[{i}]/" : $"{prefix}{child.Name.LocalName}/";
                    FlattenElement(child, newRow, rows, childPrefix);
                }
                return;
            }
        }
        rows.Add(currentRow);
    }

    /// <summary>
    /// Converte um arquivo Excel (XLSX) para XML de forma assíncrona.
    /// </summary>
    /// <param name="nome"></param>
    /// <param name="base64"></param>
    /// <returns></returns>
    public static async Task<string> ExcelToXml(string nome, string base64)
    {
        if (string.IsNullOrEmpty(nome) || nome.ToLower().Contains(".xlsx") || string.IsNullOrEmpty(base64))
        {
            return "Arquivo Excel não encontrado.";
        }


        using var workbook = new XLWorkbook(new MemoryStream(Convert.FromBase64String(base64)));
        var worksheet = workbook.Worksheets.First();

        var headers = worksheet.Row(1).Cells().Select(c => c.GetString()).ToList();
        var rows = worksheet.RowsUsed().Skip(1);

        var root = new XElement("Registros");

        foreach (var row in rows)
        {
            var record = new XElement("Registro");
            for (int i = 0; i < headers.Count; i++)
            {
                var header = headers[i];
                var value = row.Cell(i + 1).GetString();
                record.Add(new XElement(header, value));
            }
            root.Add(record);
        }

        // Create a memory stream to hold the XML output
        var outputXmlMemory = new MemoryStream();
        // Save the XML document to the memory stream
        var doc = new XDocument(root);
        // Save the XML document to the memory stream
        await doc.SaveAsync(outputXmlMemory, System.Xml.Linq.SaveOptions.None, CancellationToken.None);
        // Reset the position to the beginning of the stream before converting to Base64
        outputXmlMemory.Position = 0;
        // Convert the memory stream to a Base64 string
        await Task.CompletedTask;
        return Convert.ToBase64String(outputXmlMemory.ToArray());
    }
}