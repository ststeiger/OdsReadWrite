
namespace TestOdsReaderWriter
{


    class Usage
    {


        public static void Test()
        {
            string fileToWrite = @"myfile.ods";
            string controlFile = @"myfile_what_was_read.ods";

            System.Data.DataSet dsOut = GenerateDataSet(3, 10, 100);


            
            new libOdsReadWrite.OdsReaderWriter().WriteOdsFile(dsOut, fileToWrite);

            System.Data.DataSet dsRead = new libOdsReadWrite.OdsReaderWriter().ReadOdsFile(fileToWrite);

            new libOdsReadWrite.OdsReaderWriter().WriteOdsFile(dsRead, controlFile);
        } // End Sub Test 


        public static System.Data.DataSet GenerateDataSet(int sheetsCount, int columnsCount, int rowsCount)
        {
            System.Data.DataSet data = new System.Data.DataSet();

            for (int i = 0; i < sheetsCount; ++i)
            {
                System.Data.DataTable table = data.Tables.Add(string.Format(System.Globalization.CultureInfo.InvariantCulture, "Sheet {0}", i + 1));
                table.TableName = "Table " + (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);

                for (int j = 0; j < columnsCount; ++j)
                {
                    table.Columns.Add(string.Format(System.Globalization.CultureInfo.InvariantCulture, "Column {0}", j + 1), typeof(string));
                } // Next j 

                for (int j = 0; j < rowsCount; ++j)
                {
                    string[] columnValues = new string[columnsCount];

                    for (int k = 0; k < columnValues.Length; ++k)
                    {
                        columnValues[k] = string.Format(System.Globalization.CultureInfo.InvariantCulture, "Sheet {0} Row {1} Column {2}", i + 1, j + 1, k + 1);
                    } // Next k 

                    table.Rows.Add(columnValues);
                } // Next j 
                    
            } // Next i 

            return data;
        } // End Function GenerateDataSet 


    } // End Class Usage 


} // End Namespace TestOdsReaderWriter 
