using ExcelDataReader;


namespace ArborProposalToPdf
{
  public class ExcelReader
  {
    public ExcelReader()
    {
      System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }
    public List<string> ReadLinks(string path)
    {
      List<string> links = new();
      using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
      {
        // Auto-detect format, supports:
        //  - Binary Excel files (2.0-2003 format; *.xls)
        //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
          // Choose one of either 1 or 2:

          // 1. Use the reader methods
          do
          {
            while (reader.Read())
            {
              //Console.WriteLine()
            }
          } while (reader.NextResult());

          // 2. Use the AsDataSet extension method
          //var result = reader.AsDataSet();

          // The result of each spreadsheet is in result.Tables

          //foreach (var cell in workSheet["E2:E552"])
          //{
          //  Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
          //  links.Add(cell.Text);
          //}
        }
      }
      return links;
    }
  }
}
