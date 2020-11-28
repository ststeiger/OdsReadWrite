
namespace libOdsReadWrite
{


    // https://www.codeproject.com/Articles/38425/How-to-Read-and-Write-ODF-ODS-Files-OpenDocument-2
    public sealed class OdsReaderWriter
    {
        // Namespaces. We need this to initialize XmlNamespaceManager so that we can search XmlDocument.
        private static string[,] namespaces = new string[,]
        {
            {"table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"},
            {"office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"},
            {"style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"},
            {"text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"},
            {"draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"},
            {"fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"},
            {"dc", "http://purl.org/dc/elements/1.1/"},
            {"meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"},
            {"number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"},
            {"presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"},
            {"svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"},
            {"chart", "urn:oasis:names:tc:opendocument:xmlns:chart:1.0"},
            {"dr3d", "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"},
            {"math", "http://www.w3.org/1998/Math/MathML"},
            {"form", "urn:oasis:names:tc:opendocument:xmlns:form:1.0"},
            {"script", "urn:oasis:names:tc:opendocument:xmlns:script:1.0"},
            {"ooo", "http://openoffice.org/2004/office"},
            {"ooow", "http://openoffice.org/2004/writer"},
            {"oooc", "http://openoffice.org/2004/calc"},
            {"dom", "http://www.w3.org/2001/xml-events"},
            {"xforms", "http://www.w3.org/2002/xforms"},
            {"xsd", "http://www.w3.org/2001/XMLSchema"},
            {"xsi", "http://www.w3.org/2001/XMLSchema-instance"},
            {"rpt", "http://openoffice.org/2005/report"},
            {"of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"},
            {"rdfa", "http://docs.oasis-open.org/opendocument/meta/rdfa#"},
            {"config", "urn:oasis:names:tc:opendocument:xmlns:config:1.0"}
        };

        // Read zip stream (.ods file is zip file).
        private Ionic.Zip.ZipFile GetZipFile(System.IO.Stream stream)
        {
            return Ionic.Zip.ZipFile.Read(stream);
        }

        // Read zip file (.ods file is zip file).
        private Ionic.Zip.ZipFile GetZipFile(string inputFilePath)
        {
            return Ionic.Zip.ZipFile.Read(inputFilePath);
        }

        private System.Xml.XmlDocument GetContentXmlFile(Ionic.Zip.ZipFile zipFile)
        {
            // Get file(in zip archive) that contains data ("content.xml").
            Ionic.Zip.ZipEntry contentZipEntry = zipFile["content.xml"];

            // Extract that file to MemoryStream.
            System.IO.Stream contentStream = new System.IO.MemoryStream();
            contentZipEntry.Extract(contentStream);
            contentStream.Seek(0, System.IO.SeekOrigin.Begin);

            // Create XmlDocument from MemoryStream (MemoryStream contains content.xml).
            System.Xml.XmlDocument contentXml = new System.Xml.XmlDocument();
            contentXml.Load(contentStream);

            return contentXml;
        }

        private System.Xml.XmlNamespaceManager InitializeXmlNamespaceManager(System.Xml.XmlDocument xmlDocument)
        {
            System.Xml.XmlNamespaceManager nmsManager = new System.Xml.XmlNamespaceManager(xmlDocument.NameTable);

            for (int i = 0; i < namespaces.GetLength(0); i++)
                nmsManager.AddNamespace(namespaces[i, 0], namespaces[i, 1]);

            return nmsManager;
        }

        /// <summary>
        /// Read .ods file and store it in DataSet.
        /// </summary>
        /// <param name="inputFilePath">Path to the .ods file.</param>
        /// <returns>DataSet that represents .ods file.</returns>
        public System.Data.DataSet ReadOdsFile(string inputFilePath)
        {
            Ionic.Zip.ZipFile odsZipFile = this.GetZipFile(inputFilePath);

            // Get content.xml file
            System.Xml.XmlDocument contentXml = this.GetContentXmlFile(odsZipFile);

            // Initialize XmlNamespaceManager
            System.Xml.XmlNamespaceManager nmsManager = this.InitializeXmlNamespaceManager(contentXml);

            System.Data.DataSet odsFile = new System.Data.DataSet(System.IO.Path.GetFileName(inputFilePath));

            foreach (System.Xml.XmlNode tableNode in this.GetTableNodes(contentXml, nmsManager))
                odsFile.Tables.Add(this.GetSheet(tableNode, nmsManager));

            return odsFile;
        }

        // In ODF sheet is stored in table:table node
        private System.Xml.XmlNodeList GetTableNodes(System.Xml.XmlDocument contentXmlDocument, System.Xml.XmlNamespaceManager nmsManager)
        {
            return contentXmlDocument.SelectNodes("/office:document-content/office:body/office:spreadsheet/table:table", nmsManager);
        }

        private System.Data.DataTable GetSheet(System.Xml.XmlNode tableNode, System.Xml.XmlNamespaceManager nmsManager)
        {
            System.Data.DataTable sheet = new System.Data.DataTable(tableNode.Attributes["table:name"].Value);

            System.Xml.XmlNodeList rowNodes = tableNode.SelectNodes("table:table-row", nmsManager);

            int rowIndex = 0;
            foreach (System.Xml.XmlNode rowNode in rowNodes)
                this.GetRow(rowNode, sheet, nmsManager, ref rowIndex);

            return sheet;
        }

        private void GetRow(System.Xml.XmlNode rowNode, System.Data.DataTable sheet, System.Xml.XmlNamespaceManager nmsManager, ref int rowIndex)
        {
            System.Xml.XmlAttribute rowsRepeated = rowNode.Attributes["table:number-rows-repeated"];
            if (rowsRepeated == null || System.Convert.ToInt32(rowsRepeated.Value, System.Globalization.CultureInfo.InvariantCulture) == 1)
            {
                while (sheet.Rows.Count < rowIndex)
                    sheet.Rows.Add(sheet.NewRow());

                System.Data.DataRow row = sheet.NewRow();

                System.Xml.XmlNodeList cellNodes = rowNode.SelectNodes("table:table-cell", nmsManager);

                int cellIndex = 0;
                foreach (System.Xml.XmlNode cellNode in cellNodes)
                    this.GetCell(cellNode, row, nmsManager, ref cellIndex);

                sheet.Rows.Add(row);

                rowIndex++;
            }
            else
            {
                rowIndex += System.Convert.ToInt32(rowsRepeated.Value, System.Globalization.CultureInfo.InvariantCulture);
            }

            // sheet must have at least one cell
            if (sheet.Rows.Count == 0)
            {
                sheet.Rows.Add(sheet.NewRow());
                sheet.Columns.Add();
            }
        }

        private void GetCell(System.Xml.XmlNode cellNode, System.Data.DataRow row, System.Xml.XmlNamespaceManager nmsManager, ref int cellIndex)
        {
            System.Xml.XmlAttribute cellRepeated = cellNode.Attributes["table:number-columns-repeated"];

            if (cellRepeated == null)
            {
                System.Data.DataTable sheet = row.Table;

                while (sheet.Columns.Count <= cellIndex)
                    sheet.Columns.Add();

                row[cellIndex] = this.ReadCellValue(cellNode);

                cellIndex++;
            }
            else
            {
                cellIndex += System.Convert.ToInt32(cellRepeated.Value, System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        private string ReadCellValue(System.Xml.XmlNode cell)
        {
            System.Xml.XmlAttribute cellVal = cell.Attributes["office:value"];

            if (cellVal == null)
                return string.IsNullOrEmpty(cell.InnerText) ? null : cell.InnerText;
            else
                return cellVal.Value;
        }

        /// <summary>
        /// Writes DataSet as .ods file.
        /// </summary>
        /// <param name="odsFile">DataSet that represent .ods file.</param>
        /// <param name="outputFilePath">The name of the file to save to.</param>
        public void WriteOdsFile(System.Data.DataSet odsFile, string outputFilePath)
        {
            Ionic.Zip.ZipFile templateFile = this.GetZipFile(System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(typeof(OdsReaderWriter).Namespace+ ".template.ods"));

            System.Xml.XmlDocument contentXml = this.GetContentXmlFile(templateFile);

            System.Xml.XmlNamespaceManager nmsManager = this.InitializeXmlNamespaceManager(contentXml);

            System.Xml.XmlNode sheetsRootNode = this.GetSheetsRootNodeAndRemoveChildrens(contentXml, nmsManager);

            foreach (System.Data.DataTable sheet in odsFile.Tables)
                this.SaveSheet(sheet, sheetsRootNode);

            this.SaveContentXml(templateFile, contentXml);

            templateFile.Save(outputFilePath);
        }

        private System.Xml.XmlNode GetSheetsRootNodeAndRemoveChildrens(System.Xml.XmlDocument contentXml, System.Xml.XmlNamespaceManager nmsManager)
        {
            System.Xml.XmlNodeList tableNodes = this.GetTableNodes(contentXml, nmsManager);

            System.Xml.XmlNode sheetsRootNode = tableNodes.Item(0).ParentNode;
            // remove sheets from template file
            foreach (System.Xml.XmlNode tableNode in tableNodes)
                sheetsRootNode.RemoveChild(tableNode);

            return sheetsRootNode;
        }

        private void SaveSheet(System.Data.DataTable sheet, System.Xml.XmlNode sheetsRootNode)
        {
            System.Xml.XmlDocument ownerDocument = sheetsRootNode.OwnerDocument;

            System.Xml.XmlNode sheetNode = ownerDocument.CreateElement("table:table", this.GetNamespaceUri("table"));

            System.Xml.XmlAttribute sheetName = ownerDocument.CreateAttribute("table:name", this.GetNamespaceUri("table"));
            sheetName.Value = sheet.TableName;
            sheetNode.Attributes.Append(sheetName);

            this.SaveColumnDefinition(sheet, sheetNode, ownerDocument);

            this.SaveRows(sheet, sheetNode, ownerDocument);

            sheetsRootNode.AppendChild(sheetNode);
        }

        private void SaveColumnDefinition(System.Data.DataTable sheet, System.Xml.XmlNode sheetNode, System.Xml.XmlDocument ownerDocument)
        {
            System.Xml.XmlNode columnDefinition = ownerDocument.CreateElement("table:table-column", this.GetNamespaceUri("table"));

            System.Xml.XmlAttribute columnsCount = ownerDocument.CreateAttribute("table:number-columns-repeated", this.GetNamespaceUri("table"));
            columnsCount.Value = sheet.Columns.Count.ToString(System.Globalization.CultureInfo.InvariantCulture);
            columnDefinition.Attributes.Append(columnsCount);

            sheetNode.AppendChild(columnDefinition);
        }

        private void SaveRows(System.Data.DataTable sheet, System.Xml.XmlNode sheetNode, System.Xml.XmlDocument ownerDocument)
        {
            System.Data.DataRowCollection rows = sheet.Rows;
            for (int i = 0; i < rows.Count; i++)
            {
                System.Xml.XmlNode rowNode = ownerDocument.CreateElement("table:table-row", this.GetNamespaceUri("table"));

                this.SaveCell(rows[i], rowNode, ownerDocument);

                sheetNode.AppendChild(rowNode);
            }
        }

        private void SaveCell(System.Data.DataRow row, System.Xml.XmlNode rowNode, System.Xml.XmlDocument ownerDocument)
        {
            object[] cells = row.ItemArray;

            for (int i = 0; i < cells.Length; i++)
            {
                System.Xml.XmlElement cellNode = ownerDocument.CreateElement("table:table-cell", this.GetNamespaceUri("table"));

                if (row[i] != System.DBNull.Value)
                {
                    // We save values as text (string)
                    System.Xml.XmlAttribute valueType = ownerDocument.CreateAttribute("office:value-type", this.GetNamespaceUri("office"));
                    valueType.Value = "string";
                    cellNode.Attributes.Append(valueType);

                    System.Xml.XmlElement cellValue = ownerDocument.CreateElement("text:p", this.GetNamespaceUri("text"));
                    cellValue.InnerText = row[i].ToString();
                    cellNode.AppendChild(cellValue);
                }

                rowNode.AppendChild(cellNode);
            }
        }

        private void SaveContentXml(Ionic.Zip.ZipFile templateFile, System.Xml.XmlDocument contentXml)
        {
            templateFile.RemoveEntry("content.xml");

            System.IO.MemoryStream memStream = new System.IO.MemoryStream();
            contentXml.Save(memStream);
            memStream.Seek(0, System.IO.SeekOrigin.Begin);

            templateFile.AddEntry("content.xml", memStream);
        }

        private string GetNamespaceUri(string prefix)
        {
            for (int i = 0; i < namespaces.GetLength(0); i++)
            {
                if (namespaces[i, 0] == prefix)
                    return namespaces[i, 1];
            }

            throw new System.InvalidOperationException("Can't find that namespace URI");
        }
    }
}
