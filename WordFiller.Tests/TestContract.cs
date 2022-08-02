using System;

namespace WordFiller.Tests
{
    public class TestContract
    {
        [WordPlaceHolder("Property", NullValueHandling.ReplaceWithEmptyString)]
        public DateTime Property { get; set; }

        [WordPlaceHolder("Property1", NullValueHandling.DontReplace)]
        public int? Property1 { get; set; }

        [WordPlaceHolder("Property2", NullValueHandling.ReplaceWithEmptyString)]
        public string Property2 { get; set; }

        [WordPlaceHolder("DifferentPlaceHolder", NullValueHandling.ThrowException)]
        public double? Property3 { get; set; }
    }
}
