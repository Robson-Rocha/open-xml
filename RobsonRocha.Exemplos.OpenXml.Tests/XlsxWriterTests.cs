namespace RobsonRocha.Exemplos.OpenXml.Tests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;

    [TestClass]
    public class XlsxWriterTests
    {
        [TestMethod]
        public void WriteXlsxTest()
        {
            XlsxWriter writer = new XlsxWriter();
            string xlsxOutputPath = $@".\assets\_output_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
            writer.WriteXlsx(@".\assets\template.xlsx", xlsxOutputPath, new[] {
                new WriteXlsxOptions
                {
                    SheetName = "Sheet1",
                    IndividualReplacements = new Dictionary<string, object> {
                        { "##NAME##", "Lorem Ipsum" },
                        { "##AGE##", 35 },
                        { "##BIRTHDATE##", new DateTime(1982, 1, 18) },
                        { "##HEIGHT##", 1.85M },
                        { "##EMAIL##", "Lorem@Ipsum.com" },
                        { "##ISMARRIED##", true },
                    },
                    LineLocators = new []{
                        "##SON_NAME##", "##SON_AGE##", "##SON_EMAIL##", "##SON_BIRTHDATE##", "##SON_HEIGHT##", "##SON_GENDER##"
                    },
                    LineReplacements = new []{
                        new Dictionary<string, object> {
                            { "##SON_NAME##", "Agnes" },
                            { "##SON_AGE##", 2 },
                            { "##SON_EMAIL##", "agnes@goncalvesaraujo.com.br" },
                            { "##SON_BIRTHDATE##", new DateTime(2015, 6, 17) },
                            { "##SON_HEIGHT##", 0.65 },
                            { "##SON_GENDER##", "F" }
                        },
                        new Dictionary<string, object> {
                            { "##SON_NAME##", "Bernardo" },
                            { "##SON_AGE##", 10 },
                            { "##SON_EMAIL##", "bernardo@goncalvesaraujo.com.br" },
                            { "##SON_BIRTHDATE##", new DateTime(2007, 10, 20) },
                            { "##SON_HEIGHT##", 1.50 },
                            { "##SON_GENDER##", "M" }
                        },
                        new Dictionary<string, object> {
                            { "##SON_NAME##", "Miguel" },
                            { "##SON_AGE##", 8 },
                            { "##SON_EMAIL##", "miguel@goncalvesaraujo.com.br" },
                            { "##SON_BIRTHDATE##", new DateTime(2017, 11, 11) },
                            { "##SON_HEIGHT##", 1.25 },
                            { "##SON_GENDER##", "M" }
                        },
                    }
                }
            });
            Assert.IsTrue(File.Exists(xlsxOutputPath));
            Process.Start(xlsxOutputPath);
        }
    }
}
