using System;
using DocumentFormat.OpenXml.VariantTypes;
using FluentAssertions;
using OpenXmlDocumentLibrary;
using Xunit;

namespace OpenXmlLibrary.Tests.OpenXmlDocument
{
    public class CreateProperty
    {
        IOpenXmlDocument CreateOpenXmlDocument()
        {
            return new OpenXmlDocumentLibrary.OpenXmlDocument("test.docx");
        }

        [Fact]
        public void ShouldNeverThrow_NullReferenceException()
        {
            var openXmlDoc = CreateOpenXmlDocument();
            openXmlDoc.Invoking(x => x.CreateProperty(null, PropertyType.Text)).ShouldNotThrow<NullReferenceException>();
        }

        [Fact]
        public void ShouldThrow_WhenValueIsNull()
        {
            var openXmlDoc = CreateOpenXmlDocument();
            openXmlDoc.Invoking(x => x.CreateProperty(null, PropertyType.Text))
                .ShouldThrow<ArgumentNullException>()
                .And.ParamName.Should().Contain("propertyValue");
        }

        [Fact]
        public void ShouldReturn_DateTime()
        {
            var openXmlDoc = CreateOpenXmlDocument();

            var now = DateTime.Now;
            var property = openXmlDoc.CreateProperty(now, PropertyType.DateTime);

            property.VTFileTime.Should().NotBeNull();
        }

        [Fact]
        public void ShouldReturn_Text()
        {
            var openXmlDoc = CreateOpenXmlDocument();

            var now = DateTime.Now;
            var property = openXmlDoc.CreateProperty(now, PropertyType.Text);

            property.FirstChild.InnerText.Should().Be(now.ToString());
        }

        [Fact]
        public void ShouldReturn_true()
        {
            var openXmlDoc = CreateOpenXmlDocument();

            var property = openXmlDoc.CreateProperty(true, PropertyType.YesNo);

            property.FirstChild.InnerText.Should().Be("true");
        }


    }
}
