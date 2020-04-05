// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            var templateDoc = new FileInfo("../../TemplateDocument.docx");
            var dataFile = new FileInfo(Path.Combine(tempDi.FullName, "Data.xml"));

            // The following method generates a large data file with random data.
            // In a real world scenario, this is where you would query your data source and produce XML that will drive your document generation process.
            var data = GenerateDataFromDataSource(dataFile);

            var wmlDoc = new WmlDocument(templateDoc.FullName);
            var count = 1;
            foreach (var customer in data.Elements("Customer"))
            {
                var assembledDoc = new FileInfo(Path.Combine(tempDi.FullName, string.Format("Letter-{0:0000}.docx", count++)));
                Console.WriteLine(assembledDoc.Name);
                var wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, customer, out var templateError);
                if (templateError)
                {
                    Console.WriteLine("Errors in template.");
                    Console.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
                }
                wmlAssembledDoc.SaveAs(assembledDoc.FullName);
            }
        }

        private static readonly string[] s_productNames = new[] {
            "Unicycle",
            "Bicycle",
            "Tricycle",
            "Skateboard",
            "Roller Blades",
            "Hang Glider",
        };

        private static XElement GenerateDataFromDataSource(FileInfo dataFi)
        {
            var numberOfDocumentsToGenerate = 500;
            var customers = new XElement("Customers");
            var r = new Random();
            for (var i = 0; i < numberOfDocumentsToGenerate; ++i)
            {
                var customer = new XElement("Customer",
                    new XElement("CustomerID", i + 1),
                    new XElement("Name", "Eric White"),
                    new XElement("HighValueCustomer", r.Next(2) == 0 ? "True" : "False"),
                    new XElement("Orders"));
                var orders = customer.Element("Orders");
                var numberOfOrders = r.Next(10) + 1;
                for (var j = 0; j < numberOfOrders; j++)
                {
                    var order = new XElement("Order",
                        new XAttribute("Number", j + 1),
                        new XElement("ProductDescription", s_productNames[r.Next(s_productNames.Length)]),
                        new XElement("Quantity", r.Next(10)),
                        new XElement("OrderDate", "September 26, 2015"));
                    orders.Add(order);
                }
                customers.Add(customer);
            }
            customers.Save(dataFi.FullName);
            return customers;
        }
    }
}