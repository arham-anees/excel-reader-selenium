// See https://aka.ms/new-console-template for more information
using ArborProposalToPdf;

using System.ComponentModel;

var links = new ExcelReader().ReadLinks("C:\\Users\\Latitude\\Downloads\\arbor-proposals.xlsx");

foreach (var link in links)
{
  Console.WriteLine(link);
}
