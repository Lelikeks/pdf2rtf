using System.Text.RegularExpressions;

namespace pdf2rtf
{
    class TextCapture
    {
        private Regex _regex { get; set; }
        public string[] Properties { get; set; }

        public TextCapture(string regex, params string[] properties)
        {
            _regex = new Regex(regex);
            Properties = properties;
        }

        public Match Match(string text)
        {
            return _regex.Match(text);
        }
    }
}
