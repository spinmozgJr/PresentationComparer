using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace powerPoint
{
    class PresentationSeparator
    {
        public Image[] ConvertToPictures(string path, int width, int height)
        {
            var ppApp = new Microsoft.Office.Interop.PowerPoint.Application();

            ppApp.Visible = MsoTriState.msoCTrue;
            Presentations ppPresens = ppApp.Presentations;
            Presentation objPres = ppPresens.Open(path,
                MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);

            Image[] images = new Bitmap[objPres.Slides.Count];
            for (int i = 1; i <= objPres.Slides.Count; i++)
            {
                objPres.Slides[i].Export($"C:\\work\\slides\\NEWNAME{i}.png", "png",  width, height);
                images[i - 1] = new Bitmap($"C:\\work\\slides\\NEWNAME{i}.png");
            }

            objPres.Close();
            objPres = null;
            ppApp.Quit();
            ppApp = null;

            return images;
        }
    }
}
