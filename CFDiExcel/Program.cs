using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace CFDiExcel
{
    public class Program
    {
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                return;
            }

            var xmlFiles = GetXmlFiles(args);
            if (xmlFiles.Count == 0)
            {
                return;
            }

            var cfdis = xmlFiles.Select(x => DeserializeCfdi(File.ReadAllText(x))).ToList();
            var excelFilePath = GetNewFilePath(Path.GetDirectoryName(xmlFiles[0]), "resumen.xlsx");
            SaveExcelFile(cfdis, excelFilePath);
        }

        public static List<string> GetXmlFiles(string[] args)
        {
            var xmlFiles = new List<string>();
            var path = args[0];

            if (args.Length > 1)
            {
                xmlFiles.AddRange(args.Where(x => Path.GetExtension(x).Equals(".xml", StringComparison.OrdinalIgnoreCase)));
            }
            else if (File.Exists(path) && Path.GetExtension(path).Equals(".xml", StringComparison.OrdinalIgnoreCase))
            {
                xmlFiles.Add(path);
            }
            else if (Directory.Exists(path))
            {
                string[] files = Directory.GetFiles(path, "*.xml", SearchOption.AllDirectories);
                xmlFiles.AddRange(files);
            }

            return xmlFiles;
        }

        public static Comprobante DeserializeCfdi(string xmlData)
        {
            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(Comprobante));
                using (StringReader reader = new StringReader(xmlData))
                {
                    Comprobante cfdi = (Comprobante)serializer.Deserialize(reader);
                    return cfdi;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static string GetNewFilePath(string path, string fileName)
        {
            string? directory = Directory.Exists(path) ? path : Path.GetDirectoryName(path);
            if (directory == null)
            {
                throw new DirectoryNotFoundException(path);
            }
            return GetUniqueName(directory, fileName);
        }

        public static string GetUniqueName(string folderPath, string name)
        {
            string pathAndFileName = Path.Combine(folderPath, name);
            string validatedName = name;
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(pathAndFileName);
            string ext = Path.GetExtension(pathAndFileName);
            int count = 1;
            while (File.Exists(Path.Combine(folderPath, validatedName)))
            {
                validatedName = string.Format("{0}{1}{2}", fileNameWithoutExt, count++, ext);
            }
            return Path.Combine(folderPath, validatedName);
        }

        public static void SaveExcelFile(List<Comprobante> cfdis, string path)
        {
            using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Hoja 1"
                };
                sheets.Append(sheet);

                var mapper = new List<FieldCellMap>();
                mapper.Add(new FieldCellMap("Fecha", c => c.Fecha));
                mapper.Add(new FieldCellMap("Folio", c => c.Folio));
                mapper.Add(new FieldCellMap("FormaPago", c => c.FormaPago));
                mapper.Add(new FieldCellMap("LugarExpedicion", c => c.LugarExpedicion));
                mapper.Add(new FieldCellMap("MetodoPago", c => c.MetodoPago));
                mapper.Add(new FieldCellMap("Moneda", c => c.Moneda));
                mapper.Add(new FieldCellMap("NoCertificado", c => c.NoCertificado));
                mapper.Add(new FieldCellMap("Serie", c => c.Serie));
                mapper.Add(new FieldCellMap("TipoDeComprobante", c => c.TipoDeComprobante));
                mapper.Add(new FieldCellMap("SubTotal", c => c.SubTotal));
                mapper.Add(new FieldCellMap("Total", c => c.Total));

                mapper.Add(new FieldCellMap("EmisorNombre", c => c.Emisor.Nombre));
                mapper.Add(new FieldCellMap("EmisorRegimenFiscal", c => c.Emisor.RegimenFiscal));
                mapper.Add(new FieldCellMap("EmisorRFC", c => c.Emisor.Rfc));

                mapper.Add(new FieldCellMap("ReceptorNombre", c => c.Receptor.Nombre));
                mapper.Add(new FieldCellMap("ReceptorRegimenFiscal", c => c.Receptor.RegimenFiscalReceptor));
                mapper.Add(new FieldCellMap("ReceptorRFC", c => c.Receptor.Rfc));
                mapper.Add(new FieldCellMap("ReceptorUsoCFDI", c => c.Receptor.UsoCFDI));
                mapper.Add(new FieldCellMap("ReceptorDomicilioFiscal", c => c.Receptor.DomicilioFiscalReceptor));

                mapper.Add(new FieldCellMap("TrasladoISR", c => c.Impuestos?.Traslados?.FirstOrDefault(x => x.Impuesto == "001")?.Importe ?? ""));
                mapper.Add(new FieldCellMap("TrasladoIVA", c => c.Impuestos?.Traslados?.FirstOrDefault(x => x.Impuesto == "002")?.Importe ?? ""));
                mapper.Add(new FieldCellMap("TrasladoIEPS", c => c.Impuestos?.Traslados?.FirstOrDefault(x => x.Impuesto == "003")?.Importe ?? ""));
                mapper.Add(new FieldCellMap("RetencionISR", c => c.Impuestos?.Retenciones?.FirstOrDefault(x => x.Impuesto == "001")?.Importe ?? ""));
                mapper.Add(new FieldCellMap("RetencionIVA", c => c.Impuestos?.Retenciones?.FirstOrDefault(x => x.Impuesto == "002")?.Importe ?? ""));
                mapper.Add(new FieldCellMap("RetencionIEPS", c => c.Impuestos?.Retenciones?.FirstOrDefault(x => x.Impuesto == "003")?.Importe ?? ""));

                // Agregar datos al archivo Excel
                Row headerRow = new Row();
                foreach (var field in mapper)
                {
                    headerRow.Append(
                        new Cell() { CellValue = new CellValue(field.Name), DataType = field.DataType }
                    );
                }
                sheetData.AppendChild(headerRow);

                foreach (var cfdi in cfdis)
                {
                    Row dataRow = new Row();
                    foreach (var field in mapper)
                    {
                        dataRow.Append(
                            new Cell() { CellValue = new CellValue(field.Field(cfdi)), DataType = CellValues.String }
                        );
                    }
                    sheetData.AppendChild(dataRow);
                }
                workbookPart.Workbook.Save();
            }

            Console.WriteLine(path);
        }

        public class FieldCellMap
        {
            public FieldCellMap(string name, Func<Comprobante, string> field, EnumValue<CellValues> dataType = null)
            {
                Name = name ?? throw new ArgumentNullException(nameof(name));
                Field = field ?? throw new ArgumentNullException(nameof(field));
                DataType = dataType ?? CellValues.String;
            }

            public string Name { get; set; }

            public Func<Comprobante, string> Field { get; set; }

            public EnumValue<CellValues> DataType { get; set; }
        }
    }
}
