// See https://aka.ms/new-console-template for more information
using ArborProposalToPdf;


var links = new ExcelReader().ReadLinks("C:\\Users\\Latitude\\Downloads\\arbor-proposals.xlsx");

Console.WriteLine($"total links {links.Count}");
var browserSupport = new BrowserSupport();
int count = 0;
foreach (var link in links)
{
  if (count > 3) break;
  browserSupport.OpenBrowser(link); 
    count++;
}

