using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocxTemplateParser;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DocxTemplateTester.Tests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            var file = @"C:\Users\nertil\Downloads\ModifyImage (1).docx";
            var parser = new TemplateParser();
            var fileInput = File.Open(file, FileMode.OpenOrCreate);
            var data = new
            {
                FirstName = "Nertil",
                LastName = "Poci",
                Person = new
                {
                    FirstName = "Nertil",
                    LastName = "Poci"
                },
                Persons = new List<object> {
new
                {
                    FirstName = "Nertil1",
                    LastName = "Poci1"
                },
new
                {
                    FirstName = "Nertil2",
                    LastName = "Poci2"
                },
new
                {
                    FirstName = "Nertil3",
                    LastName = "Poci3"
                }
                }.AsQueryable()
            };
            parser.Parse(fileInput,data);
        }
    }
}
