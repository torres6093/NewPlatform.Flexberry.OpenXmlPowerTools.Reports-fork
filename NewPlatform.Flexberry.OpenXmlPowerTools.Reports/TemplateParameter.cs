namespace NewPlatform.Flexberry.Reports
{
    /// <summary>
    /// Параметр в шаблоне
    /// <example>
    /// [Договор.ДатаВПоземельнойКниге:dd.MM.yyyy]
    /// </example>
    /// </summary>
    public class TemplateParameter
    {
        private string format;

        public TemplateParameter(string fullName)
        {
            FullName = fullName;

            var colonIndex = fullName.IndexOf(":");
            string paramName, paramFormat = string.Empty;
            if (colonIndex > 0)
            {
                paramName = fullName.Substring(0, colonIndex);
                paramFormat = fullName.Substring(colonIndex + 1);
            }
            else
            {
                paramName = fullName;
            }

            Name = paramName;
            format = "{0";
            if (!string.IsNullOrEmpty(paramFormat))
            {
                format += ":" + paramFormat;
            }

            format += "}";
        }

        public string Name { get; private set; }

        public string FullName { get; private set; }

        internal string FormatObject(object dataObject)
        {
            return string.Format(format, dataObject);
        }
    }
}
