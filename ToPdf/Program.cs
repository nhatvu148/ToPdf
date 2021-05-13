using OfficeConverter;
using System;
using System.IO;
using System.Reflection;

namespace ToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            string pathDirectory = $"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}";
            string docPath = $"{pathDirectory}/../対象船のデータ取得状況.docx";

            if (args == null || args.Length == 0)
            {
                // no arguments
            }
            else
            {
                docPath = Convert.ToString(args[0]);
            }

            Console.WriteLine($"Start converting {docPath} to pdf...");
            using (var converter = new Converter())
            {
                converter.Convert(docPath, docPath.Replace(".docx", ".pdf"));
            }
            Console.WriteLine("Finished!");
        }
    }
}
