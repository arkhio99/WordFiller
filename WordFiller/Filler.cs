using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using FastMember;
using OpenXmlPowerTools;

namespace WordFiller
{
    public class Filler<T>
    {
        private Type _usingType;
        private TypeAccessor _accessor;
        private List<(string PlaceholderName, string PropertyName, NullValueHandling NullValueHandling)> _propertyData;

        public Filler()
        {
            _usingType = typeof(T);
            _accessor = TypeAccessor.Create(_usingType);
            _propertyData = _usingType.GetProperties()
                .Select(p => (Attributes: p.GetCustomAttributes(typeof(WordPlaceHolderAttribute), false), Name: p.Name))
                .Where(ap => ap.Attributes.Length > 0)
                .Select(ap => (
                    PlaceholderName: ((WordPlaceHolderAttribute)ap.Attributes[0]).Name,
                    PropertyName: ap.Name,
                    NullValueHandling: ((WordPlaceHolderAttribute)ap.Attributes[0]).NullValueHandling))
                .ToList();
        }

        public WordprocessingDocument FillDocument(WordprocessingDocument document, T data)
        {
            ValidateData(data);

            var replacements = _propertyData
                .Where(tuple => !(tuple.NullValueHandling == NullValueHandling.DontReplace && _accessor[data, tuple.PropertyName] == null))
                .Select(
                tuple => (
                    PlaceHolder: tuple.PlaceholderName,
                    Replacement: _accessor[data, tuple.PropertyName]?.ToString() ?? ""));

            Replace(document, replacements);

            return document;
        }

        private void ValidateData(T data)
        {
            var toThrowException = _propertyData
                .Where(v => v.NullValueHandling == NullValueHandling.ThrowException)
                .Select(v => v.PropertyName);

            var unfilledProperties = new List<string>();

            foreach (var propertyName in toThrowException)
            {
                if (_accessor[data, propertyName] is null)
                {
                    unfilledProperties.Add(propertyName);
                }
            }

            if (unfilledProperties.Count > 0)
            {
                throw new ArgumentException($"Properties [{string.Join(", ", unfilledProperties)}] have to be filled, but they are not.");
            }
        }

        private void Replace(WordprocessingDocument file, IEnumerable<(string Placeholder, string Replacement)> replacements)
        {
            var replacementByRegex = replacements
                .OrderByDescending(r => r.Placeholder.Length)
                .Select(r => (Regex: new Regex(@"\$" + r.Placeholder + @"\$"), Replacement: r.Replacement));
            var xDocument = file.MainDocumentPart.GetXDocument();
            var xElements = xDocument.Descendants(W.p);

            foreach (var rbr in replacementByRegex)
            {
                var replaced = 0;
                do
                {
                    replaced = OpenXmlRegex.Replace(xElements, rbr.Regex, rbr.Replacement, null);
                }
                while (replaced != 0);
            }

            file.MainDocumentPart.PutXDocument();
        }
    }
}
