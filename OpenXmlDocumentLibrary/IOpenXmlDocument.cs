namespace OpenXmlDocumentLibrary
{
    public interface IOpenXmlDocument
    {

        string SetNewProperty(string fileName, string propertyName, object propertyValue, PropertyType propertyType);

    }
}
