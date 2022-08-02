using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace WordFiller.Tests
{
    [TestClass]
    public class FillerTests
    {
        private static Filler<TestContract> _filler;

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            _filler = new Filler<TestContract>();
        }

        [TestMethod]
        public void Filler_FillDocument_CorrectContract_NewDocumentFilledCorrectly()
        {
            using var document = WordprocessingDocument.CreateFromTemplate("../../../Template.docx");
            var data = new TestContract
            {
                Property = new DateTime(2022, 1, 1),
                Property1 = 1,
                Property2 = "uniq string",
                Property3 = 3.1,
            };

            _filler.FillDocument(document, data);

            using var actual = new StreamReader(document.MainDocumentPart.GetStream(FileMode.Open));
            var actualXML = actual.ReadToEnd();

            using var expected = new StreamReader(
                WordprocessingDocument.Open("../../../FullyFilledExpected.docx", false)
                .MainDocumentPart
                .GetStream(FileMode.Open));
            var expectedXML = expected.ReadToEnd();

            Assert.AreEqual(actualXML, expectedXML);
        }

        [TestMethod]
        public void Filler_FillDocument_Property3Unfilled_ThrewException()
        {
            using var document = WordprocessingDocument.CreateFromTemplate("../../../Template.docx");
            var data = new TestContract
            {
                Property = new DateTime(2021, 1, 1),
                Property1 = 1,
                Property2 = "uniq string",
                Property3 = null,
            };

            Assert.ThrowsException<ArgumentException>(() => _filler.FillDocument(document, data));
        }

        [TestMethod]
        public void Filler_FillDocument_Properties1and2Unfilled_NewDOcumentFilledCorrectly()
        {
            using var document = WordprocessingDocument.CreateFromTemplate("../../../Template.docx");
            var data = new TestContract
            {
                Property = new DateTime(),
                Property1 = null,
                Property2 = null,
                Property3 = 3.1,
            };

            _filler.FillDocument(document, data);

            using var actual = new StreamReader(document.MainDocumentPart.GetStream(FileMode.Open));
            var actualXML = actual.ReadToEnd();

            using var expected = new StreamReader(
                WordprocessingDocument.Open("../../../NotFullyFilledExpected.docx", false)
                .MainDocumentPart
                .GetStream(FileMode.Open));
            var expectedXML = expected.ReadToEnd();

            Assert.AreEqual(actualXML, expectedXML);
        }
    }
}
