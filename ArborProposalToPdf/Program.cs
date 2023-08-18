// See https://aka.ms/new-console-template for more information
using ArborProposalToPdf;


var proposals = new ExcelReader().ReadLinks("C:\\Users\\Latitude\\Downloads\\arbor-proposals.xlsx");

Console.WriteLine($"total links {proposals.Count}");
var browserSupport = new BrowserSupport();
int count = 0;
int startFrom = 291;
List<string> failedGroups = new();
browserSupport.OpenBrowser();
foreach (var proposal in proposals)
{
  
  Console.WriteLine($"Proposal #{count + 1}");
  Console.WriteLine(proposal.Link);
  count++;
  if (startFrom <= count)
  {
    try
    {
      browserSupport.Navigate(proposal.Link);
      browserSupport.Print($"{proposal.GroupId}.pdf");
    }
    catch(Exception e)
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


