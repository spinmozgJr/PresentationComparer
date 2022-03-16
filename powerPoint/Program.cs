using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Drawing;

namespace powerPoint
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = "";
            PresentationSeparator presentationSeparator = new PresentationSeparator();
            presentationSeparator.ConvertToPictures(path);
        }
    }
}
