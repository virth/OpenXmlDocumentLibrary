using DocumentFormat.OpenXml.CustomProperties;

namespace OpenXmlDocumentLibrary
{
    public interface IOpenXmlDocument
    {

        string SetNewProperty(string propertyName, object propertyValue, PropertyType propertyType);
        CustomDocumentProperty CreatePropertyFromPropertyType(object propertyValue, PropertyType propertyType)

    }
}
