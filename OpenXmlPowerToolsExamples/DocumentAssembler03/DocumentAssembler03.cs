// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace OpenXmlPowerTools
{
    internal class Program
    {
        private static readonly string[] ProductNames =
        {
            "Unicycle",
            "Bicycle",
            "Tricycle",
            "Skateboard",
            "Roller Blades",
            "Hang Glider"
        };

        private static void Main()
        {
            DateTime n = DateTime.Now;
            var tempDi = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            tempDi.Create();

            var templateDoc = new FileInfo("../../../TemplateDocument.docx");
            var dataFile = new FileInfo(Path.Combine(tempDi.FullName, "Data.xml"));

            // The following method generates a large data file with random data.
            // In a real world scenario, this is where you would query your data source and produce XML that will drive your document generation process.
            XElement data = GenerateDataFromDataSource(dataFile);

            var wmlDoc = new WmlDocument(templateDoc.FullName);
            var count = 1;

            foreach (XElement customer in data.Elements("Customer"))
            {
                var assembledDoc = new FileInfo(Path.Combine(tempDi.FullName, $"Letter-{count++:0000}.docx"));
                Console.WriteLine(assembledDoc.Name);
                WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, customer, out bool templateError);
                if (templateError)
                {
                    Console.WriteLine("Errors in template.");
                    Console.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
                }

                wmlAssembledDoc.SaveAs(assembledDoc.FullName);
            }
        }

        private static XElement GenerateDataFromDataSource(FileInfo dataFi)
        {
            const int numberOfDocumentsToGenerate = 500;
            var customers = new XElement("Customers");
            var r = new Random();

            for (var i = 0; i < numberOfDocumentsToGenerate; ++i)
            {
                var customer = new XElement("Customer",
                    new XElement("CustomerID", i + 1),
                    new XElement("Name", "Eric White"),
                    new XElement("HighValueCustomer", r.Next(2) == 0 ? "True" : "False"),
                    new XElement("Orders"));

                XElement orders = customer.Elements("Orders").First();
                int numberOfOrders = r.Next(10) + 1;

                for (var j = 0; j < numberOfOrders; j++)
                {
                    var order = new XElement("Order",
                        new XAttribute("Number", j + 1),
                        new XElement("ProductDescription", ProductNames[r.Next(ProductNames.Length)]),
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
