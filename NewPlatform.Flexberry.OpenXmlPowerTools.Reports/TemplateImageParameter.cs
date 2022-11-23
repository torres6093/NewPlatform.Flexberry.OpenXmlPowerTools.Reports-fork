namespace NewPlatform.Flexberry.Reports
{
    public class TemplateImageParameter
    {
        public const string ImgBookmarkPrefix = "imgTemplate";

        public string Name { get; private set; }
        public string FullName { get; private set; }

        public TemplateImageParameter(string fullName)
        {
            FullName = fullName;
            Name = FullName.Replace(ImgBookmarkPrefix, "");
        }
    }
}
