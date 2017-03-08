
using System;
using FluentAssertions;
using OpenXmlDocumentLibrary;

using Xunit;

namespace OpenXmlLibrary.Tests
{
 
    public class SetNewProperty
    {
        OpenXmlDocument CreateOpenXmlDocument()
        {
            return new OpenXmlDocument("test.docx");
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
    }
}
