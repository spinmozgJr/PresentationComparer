using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Drawing;

namespace powerPoint
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = @"C:\work\test.pptx";
            int width = 320;
            int height = 240;
            PresentationSeparator presentationSeparator = new PresentationSeparator();
            presentationSeparator.ConvertToPictures(path, width, height);
        }
    }
}
