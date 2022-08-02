using System;

namespace WordFiller
{
    public class WordPlaceHolderAttribute : Attribute
    {
        public string Name { get; private set; }

        public NullValueHandling NullValueHandling { get; private set; }

        public WordPlaceHolderAttribute(string name, NullValueHandling nullValueHandling)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            NullValueHandling = nullValueHandling;
        }
    }
}
