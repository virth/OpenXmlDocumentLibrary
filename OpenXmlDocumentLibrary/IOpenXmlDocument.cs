﻿namespace OpenXmlDocumentLibrary
{
    public interface IOpenXmlDocument
    {

        string SetNewProperty(string propertyName, object propertyValue, PropertyType propertyType);

    }
}
