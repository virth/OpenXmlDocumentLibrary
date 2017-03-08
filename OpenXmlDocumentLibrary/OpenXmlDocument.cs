using System;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.VariantTypes;

namespace OpenXmlDocumentLibrary
{
    public class OpenXmlDocument : IOpenXmlDocument

    {
        private readonly string _filename;
        public OpenXmlDocument(string filename)
        {
            if (string.IsNullOrEmpty(filename))
                throw new ArgumentNullException(nameof(filename));

            _filename = filename;
        }

        /// <summary>
        ///  Given a property name/value, and the property type, add a custom property to a document. 
        /// </summary>
        /// <returns>The function returns the original value, if it existed</returns>
        public string SetNewProperty(string propertyName, object propertyValue, PropertyType propertyType)
        {
            string originalValue = null;

            var newProperty = CreatePropertyFromPropertyType(propertyValue, propertyType);

            // Now that you've handled the parameters, start
            // working on the document.
            newProperty.FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";
            newProperty.Name = propertyName;

            originalValue = AddPropertyToDocument(originalValue, newProperty);

            return originalValue;
        }

        private string AddPropertyToDocument(string returnValue, CustomDocumentProperty newProperty)
        {
            if (_filename.EndsWith("xlsx"))
                returnValue = ProcessExcel(newProperty);
            else if (_filename.EndsWith("docx"))
                returnValue = ProcessWord(newProperty);
            else
                throw new Exception("Unknown Filetype");
            return returnValue;
        }

        /// <summary>
        /// Creates the specific CustomDocumentProperty from the given propertyType with the given value
        /// </summary>
        /// <returns>returns null if property is not set</returns>
        private CustomDocumentProperty CreatePropertyFromPropertyType(object propertyValue, PropertyType propertyType)
        {
            var newProp = new CustomDocumentProperty();
            var propSet = false;

            // Calculate the correct type:
            switch (propertyType)
            {
                case PropertyType.DateTime:
                    // Make sure you were passed a real date, 
                    // and if so, format in the correct way. 
                    // The date/time value passed in should 
                    // represent a UTC date/time.
                    if ((propertyValue) is DateTime)
                    {
                        newProp.VTFileTime = new VTFileTime($"{Convert.ToDateTime(propertyValue):s}Z");
                        propSet = true;
                    }

                    break;
                case PropertyType.NumberInteger:
                    if ((propertyValue) is int)
                    {
                        newProp.VTInt32 = new VTInt32(propertyValue.ToString());
                        propSet = true;
                    }

                    break;
                case PropertyType.NumberDouble:
                    if (propertyValue is double)
                    {
                        newProp.VTFloat = new VTFloat(propertyValue.ToString());
                        propSet = true;
                    }

                    break;
                case PropertyType.Text:
                    newProp.VTLPWSTR = new VTLPWSTR(propertyValue.ToString());
                    propSet = true;

                    break;
                case PropertyType.YesNo:
                    if (propertyValue is bool)
                    {
                        newProp.VTBool = new VTBool(Convert.ToBoolean(propertyValue).ToString().ToLower());
                        propSet = true;
                    }
                    break;

                default:
                    throw new ArgumentException($"unknown {nameof(propertyType)} [{propertyType}]");
            }

            return propSet ? newProp : null;
        }

        private string ProcessExcel(CustomDocumentProperty newProperty)
        {
            var originalValue = "";
            using (var document = SpreadsheetDocument.Open(_filename, true))
            {
                originalValue = AddProperty(newProperty, document);
            }
            return originalValue;
        }

        private string ProcessWord(CustomDocumentProperty newProperty)
        {
            var originalValue = "";
            using (var document = WordprocessingDocument.Open(_filename, true))
            {
                originalValue = AddProperty(newProperty, document);
            }
            return originalValue;
        }

        private static string AddProperty(CustomDocumentProperty newProperty, OpenXmlPackage document)
        {
            var originalValue = "";
            var customProps = GetCustomDocumentProperties(document);

            var existinProperties = customProps.Properties;
            if (existinProperties == null)
                return originalValue;

            var existingProperty = existinProperties.FirstOrDefault( p => string.Equals(((CustomDocumentProperty)p).Name.Value, newProperty.Name.Value, StringComparison.CurrentCultureIgnoreCase));

            if (existingProperty != null)
            {
                originalValue = existingProperty.InnerText;
                existingProperty.Remove();
            }

            // Append the new property, and 
            // fix up all the property ID values. 
            // The PropertyId value must start at 2.
            existinProperties.AppendChild(newProperty);
            var pid = 2;
            foreach (var openXmlElement in existinProperties)
            {
                var item = (CustomDocumentProperty)openXmlElement;
                item.PropertyId = pid++;
            }
            existinProperties.Save();
            return originalValue;
        }

        private static CustomFilePropertiesPart GetCustomDocumentProperties(OpenXmlPackage document)
        {
            CustomFilePropertiesPart customProps = null;
            if (document.GetType() == typeof(WordprocessingDocument))
            {
                var word = document as WordprocessingDocument;
                if (word != null)
                    customProps = word.CustomFilePropertiesPart;
            }
            else if (document.GetType() == typeof(SpreadsheetDocument))
            {
                var excel = document as SpreadsheetDocument;
                if (excel != null)
                    customProps = excel.CustomFilePropertiesPart;
            }

            if (customProps == null)
                customProps = InitializeCustomDocumentProperties(document);
            return customProps;
        }

        /// <summary>
        /// No custom properties? Add the part, and the collection of properties 
        /// </summary>
        private static CustomFilePropertiesPart InitializeCustomDocumentProperties(OpenXmlPackage document)
        {
            CustomFilePropertiesPart customProperties = null;
            if (document.GetType() == typeof(WordprocessingDocument))
            {
                var word = document as WordprocessingDocument;
                if (word != null)
                    customProperties = word.AddCustomFilePropertiesPart();
            }
            else if (document.GetType() == typeof(SpreadsheetDocument))
            {
                var excel = document as SpreadsheetDocument;
                if (excel != null)
                    customProperties = excel.AddCustomFilePropertiesPart();
            }
            else
                throw new ArgumentException($"unkown type of {nameof(document)} [{document.GetType()}]");

            if (customProperties == null)
                throw new Exception("customDocumentProperties could not be initialized");

            customProperties.Properties = new Properties();
            return customProperties;
        }
    }
}
