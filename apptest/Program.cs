using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            string chemin_modele = @"C:\Users\seymour\Desktop\modele.docx";
            string chemin_lc = @"C:\Users\seymour\Desktop\LC.docx";
            WordRead.CreateNewWordDocument(chemin_modele);
            WordRead.ReadDocument(chemin_modele, chemin_lc, "raison_juridique", "SARL");
            Console.ReadLine();
        }
    }
}
