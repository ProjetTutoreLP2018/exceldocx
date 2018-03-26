using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using Xceed.Words.NET;


namespace ConsoleApp1
{
    class WordRead
    {
        public static void CreateNewWordDocument(string chemin)
        {
            DocX doc = DocX.Create(chemin);
            doc.InsertParagraph("Formatted paragraphs").FontSize(15d).SpacingAfter(50d).Alignment = Alignment.center;

            var p = doc.InsertParagraph();
            p.Append("{{raison_juridique}} is a simple formatted red bold paragraph").Font(new Font("Arial"));

            doc.Save();
        }

        public static void ReadDocument(string chemin_modele, string chemin_lc, string mot_origine, string mot_voulu)
        {

            DocX doc = DocX.Load(chemin_modele);
            

            foreach (Paragraph paragraph in doc.Paragraphs)
            {
                paragraph.ReplaceText("{{"+mot_origine+"}}", mot_voulu);
            }
            Console.WriteLine(doc.Text);
            doc.SaveAs(chemin_lc);
        }

    }
}

