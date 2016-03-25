using System;
using System.IO;
using System.Text.RegularExpressions;
using Visio = Microsoft.Office.Interop.Visio;

namespace TortoiseStartCommitHookVisioExportVBA
{
    class Program
    {

        static void Main(string[] args)
        {
            // string[] affectedPaths = File.ReadAllLines(args[1]);
            string[] affectedPaths = Directory.GetFiles(args[3], "*.*", SearchOption.AllDirectories);

            Regex fileExtensionPattern = new Regex(@"^.*\.(vss|vsd|vst)$", RegexOptions.IgnoreCase);
            
            Visio.InvisibleApp app = new Visio.InvisibleApp();

            foreach (string path_ in affectedPaths)
            {
                string path = Path.GetFullPath(path_);

                if (fileExtensionPattern.IsMatch(path) && File.Exists(path))
                {
                    Visio.Document doc = app.Documents.Open(path);
                    ExportVBA(doc, Path.Combine(args[0], doc.Name));
                    doc.Close();
                }
            }
            app.Quit();
        }

        private static string GetComponentFileExtension(dynamic component)
        {
            int componentType = Convert.ToInt32(component.Type);
            switch (componentType)
            {
                case 1: return ".bas";
                case 2: return ".cls";
                case 3: return ".frm";
                default:
                    return null;
            }
        }

        private static void ExportVBA(Visio.Document doc, string path)
        {
            if (doc == null)
                return;

            var project = doc.VBProject;

            if (project == null)
                return;

            if (Directory.Exists(path))
                try
                {
                    Directory.Delete(path, true);
                }
                catch
                {
                    Console.Error.WriteLine("Source file or directory blocked!");
                    Environment.Exit(1);
                }

            Directory.CreateDirectory(path);

            ExportThisDocumentVBA(path, project);

            foreach (var component in project.VBComponents)
            {
                var fileExtension = GetComponentFileExtension(component);
                if (fileExtension == null)
                    continue;

                component.Export(Path.Combine(path, component.Name + fileExtension));
            }
        }

        private static void ExportThisDocumentVBA(string path, dynamic project)
        {
            var thisDocumentComponent = project.VBComponents["ThisDocument"];
            if (thisDocumentComponent != null)
            {
                var codeModule = thisDocumentComponent.CodeModule;
                var countOfLines = Convert.ToInt32(codeModule.CountOfLines);
                if (countOfLines > 0)
                {
                    var lines = codeModule.Lines(1, countOfLines);
                    File.WriteAllText(Path.Combine(path, "ThisDocument.bas"), lines);
                }
            }
        }
    }
}
