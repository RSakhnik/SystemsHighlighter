using System.IO;

namespace SystemsHighlighter
{
    public static class IconLoader
    {
        public static byte[] GetIcon(string resourceName)
        {
            using (var stream = typeof(IconLoader).Assembly.GetManifestResourceStream($"SystemsHighlighter.Tools.{resourceName}"))
            {
                return ReadBytes(stream);
            }
        }

        private static byte[] ReadBytes(Stream input)
        {
            var ms = new MemoryStream();
            input.CopyTo(ms);
            return ms.ToArray();
        }
    }
}
