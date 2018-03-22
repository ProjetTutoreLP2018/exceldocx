using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Xceed.Words.NET;

namespace apptest
{
    class WordRead
    {
        
        public static void CreateNewWordDocument(string chemin)
        {
            DocX doc = DocX.Create(chemin);
            doc.InsertParagraph("Formatted paragraphs").FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;

            var p = doc.InsertParagraph();
            p.Append("This is a simple formatted red bold paragraph").Font(new Font("Arial"));

            doc.Save();
        }

        public static void ReadDocument(string chemin)
        {

            DocX doc = DocX.Load(chemin);
            Console.WriteLine(doc.Text);
            doc.Save();
        }

    }
}
