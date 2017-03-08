
using System;
using FluentAssertions;
using OpenXmlDocumentLibrary;

using Xunit;

namespace OpenXmlLibrary.Tests
{
 
    public class SetNewProperty
    {
        OpenXmlDocumentLibrary.OpenXmlDocument CreateOpenXmlDocument(string filename = "test.docx")
        {
            return new OpenXmlDocumentLibrary.OpenXmlDocument(filename);
        }

        [Fact]
        public void ShouldNeverThrow_NullReferenceException()
        {
            var openXmlDoc = CreateOpenXmlDocument();
            openXmlDoc.Invoking(x => x.SetNewProperty(null, null, PropertyType.Text)).ShouldNotThrow<NullReferenceException>();
        }

        [Fact]
        public void ShouldThrow_WhenPropertyNameIsNull()
        {
            var openXmlDoc = CreateOpenXmlDocument();
            openXmlDoc.Invoking(x => x.SetNewProperty(null, null, PropertyType.Text))
                .ShouldThrow<ArgumentNullException>()
                .And.ParamName.Should().Contain("propertyName");
        }

        [Fact]
        public void ShouldThrow_WhenPropertyValueIsNull()
        {
            var openXmlDoc = CreateOpenXmlDocument();
            openXmlDoc.Invoking(x => x.SetNewProperty("test", null, PropertyType.Text))
                .ShouldThrow<ArgumentNullException>()
                .And.ParamName.Should().Contain("propertyValue");
        }

        [Fact]
        public void ShouldThrow_WhenFileTypeIsUnknown()
        {
            var openXmlDoc = CreateOpenXmlDocument("foo.bar");
            openXmlDoc.Invoking(x => x.SetNewProperty("foo", "bar", PropertyType.Text))
                .ShouldThrow<Exception>()
                .And.Message.Should().Contain("Unknown");
        }
    }
}
