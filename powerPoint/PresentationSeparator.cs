using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace powerPoint
{
    class PresentationSeparator
    {
        public Image[] ConvertToPictures(string path)
        {
            var ppApp = new Microsoft.Office.Interop.PowerPoint.Application();

            ppApp.Visible = MsoTriState.msoCTrue;
            Presentations ppPresens = ppApp.Presentations;
            Presentation objPres = ppPresens.Open(@"C:\work\test.pptx",
                MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

            Image[] images = new Bitmap[objPres.Slides.Count];
            for (int i = 1; i <= objPres.Slides.Count; i++)
            {
                objPres.Slides[i].Export($"C:\\work\\slides\\NEWNAME{i}.png", "png", 1280, 720);
                images[i - 1] = new Bitmap($"C:\\work\\slides\\NEWNAME{i}.png");
            }

            objPres.Close();
            objPres = null;
            ppApp.Quit();
            ppApp = null;

            return images;
        }

        public Image[] ConvertToCompressedPictures(string path)
        {
            var ppApp = new Microsoft.Office.Interop.PowerPoint.Application();

            ppApp.Visible = MsoTriState.msoCTrue;
            Presentations ppPresens = ppApp.Presentations;
            Presentation objPres = ppPresens.Open(@"C:\work\test.pptx",
                MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

            Image[] compressedImages = new Bitmap[objPres.Slides.Count];
            for (int i = 1; i <= objPres.Slides.Count; i++)
            {
                objPres.Slides[i].Export($"C:\\work\\slides\\comprasedNEWNAME{i}.png", "png", 320, 240);
                compressedImages[i - 1] = new Bitmap($"C:\\work\\slides\\comprasedNEWNAME{i}.png");
            }

            objPres.Close();
            objPres = null;
            ppApp.Quit();
            ppApp = null;

            return compressedImages;
        }

    }
}
