namespace NewPlatform.Flexberry.Reports
{
    public class ImageParameter
    {
        public ImageParameter(string fileName, int width, int height)
        {
            FileName = fileName;
            Height = height;
            Width = width;
        }

        public string FileName { get; private set; }
        public int Height { get; private set; }
        public int Width { get; private set; }
    }
}
