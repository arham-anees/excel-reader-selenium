// See https://aka.ms/new-console-template for more information
using ArborProposalToPdf;


var proposals = new ExcelReader().ReadLinks("C:\\Users\\Latitude\\Downloads\\arbor-proposals.xlsx");
var directoryPath = "C:\\Users\\Latitude\\Downloads";
// read all pdf files from directory
var pdfs = Directory.GetFiles(directoryPath).Where(x => x.EndsWith(".pdf")).Select(x => x.Split('\\').Last()).ToList();

//var allGroups=proposals.Select(x => x.GroupId + ".pdf").ToList();
//failedGroups = allGroups.Where(x => !pdfs.Contains(x)).ToList();

//Console.WriteLine("List of failed proposals");
//failedGroups.ForEach(Console.WriteLine);
//return;

List<Task> tasks = new();

    Console.WriteLine($"total links {proposals.Count}");
List<string> failedGroups = new();
var browserSupport = new BrowserSupport();

int count = 0;
int startFrom = 0;
int endAt = 300;
browserSupport.OpenBrowser();
foreach (var proposal in proposals)
{
  count++;
  if (startFrom <= count && endAt>=count)
  {
    try
    {
      Console.WriteLine($"Proposal #{count + 1}");
      //Console.WriteLine(proposal.Link);
      browserSupport.Navigate(proposal.Link);
      browserSupport.Print($"{proposal.GroupId}.pdf");
    }
    catch (Exception e)
    {
      failedGroups.Add(proposal.GroupId);
      Console.WriteLine("Proposal failed");
    }
  }
}
browserSupport.Close();

Console.WriteLine("List of failed proposals");
failedGroups.ForEach(Console.WriteLine);
Console.WriteLine("Process completed.");

