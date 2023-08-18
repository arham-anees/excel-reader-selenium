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
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
          do
          {
            while (reader.Read())
            {
              var link = reader.GetString(4);
              if (link != "url")
              {
                links.Add(link);
              }
            }
          } while (reader.NextResult());
        }
      }
      return links;
    }
  }
}
