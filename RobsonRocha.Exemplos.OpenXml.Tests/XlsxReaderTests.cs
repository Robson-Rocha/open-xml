namespace RobsonRocha.Exemplos.OpenXml.Tests
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Collections.Generic;
    using System.Linq;

    [TestClass]
    public class XlsxTestsTests
    {
        [TestMethod]
        public void ReadXlsxTest()
        {
            XlsxReader xlsxReader = new XlsxReader();
            string xlsxPath = @".\Assets\ImportData.xlsx";
            ReadXlsxOptions[] readXlsxOptions = new[]
            {
                new ReadXlsxOptions {
                    SheetName = "data",
                    HeaderRowIndex = 1
                }
            };

            IReadOnlyList<SheetInfo> results = xlsxReader.ReadXlsx(xlsxPath, readXlsxOptions);
            Assert.IsNotNull(results);
            Assert.IsTrue(results.Any());
            Assert.IsTrue(results[0].Name == "data");
            Assert.IsTrue(results[0].Columns[1] == "first_name");
            Assert.IsTrue(results[0].GetCell("B8").Value == "Alejandra");
            Assert.IsTrue(results[0].AllRows[0][3].Value == "email");
            Assert.IsTrue(results[0].AllRows[2]["country"].Value == "China");
            Assert.IsTrue(results[0].Rows[0][3].Value == "lvibert0@utexas.edu");
            Assert.IsTrue((results[0].Rows[0] as dynamic).email == "lvibert0@utexas.edu");

        }
    }
}
