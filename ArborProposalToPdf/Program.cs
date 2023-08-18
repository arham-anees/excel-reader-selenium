// See https://aka.ms/new-console-template for more information
using ArborProposalToPdf;


var links = new ExcelReader().ReadLinks("C:\\Users\\Latitude\\Downloads\\arbor-proposals.xlsx");

foreach (var link in links)
{
  Console.WriteLine(link);
}

Console.WriteLine($"total links {links.Count}");
