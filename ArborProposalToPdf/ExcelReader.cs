using ExcelDataReader;


namespace ArborProposalToPdf
{
  public class ExcelReader
  {
    public ExcelReader()
    {
      System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }
    public List<Proposal> ReadLinks(string path)
    {
      List<Proposal> links = new();
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
              var groupId = reader.GetDouble(1).ToString();
                links.Add(new Proposal
                {
                  GroupId= groupId,
                  Link=link
                });
              }
            }
          } while (reader.NextResult());
        }
      }
      return links;
    }
  }
}
