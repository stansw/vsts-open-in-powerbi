using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Mashup.Client.Packaging;
using Microsoft.Mashup.Host.Document;
using Microsoft.Mashup.Host.Document.Storage;
using Microsoft.Mashup.Host.Document.Storage.Metadata;
using Microsoft.Mashup.Storage.Memory;
using PowerBIProtocolHandler.Properties;

namespace PowerBIProtocolHandler
{
    public class Program
    {
        private class Options
        {
            public string Url { get; set; }
            public string Project { get; set; }
            public string QueryName { get; set; }
            public string QueryId { get; set; }
        }

        public static void Main(string[] args)
        {
            try
            {
                //args = new[] { "stasiu://Query?url=https://stansw.visualstudio.com&project=Sample1&queryName=Favourite&queryId=5c4fbdb9-509b-49cc-bf6e-e22742fbd6b1" };
                var options = ParseArgs(args);
                Main(options);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
                Environment.Exit(1);
            }
        }

        private static Options ParseArgs(String[] args)
        {
            if (args.Length != 1)
            {
                throw new InvalidOperationException("Invalid number of arguments.");
            }

            var uriText = args[0];
            var uri = default(Uri);
            if (!Uri.TryCreate(uriText, UriKind.Absolute, out uri))
            {
                throw new InvalidOperationException("Invalid format.");
            }

            var query = HttpUtility.ParseQueryString(uri.Query);

            return new Options()
            {
                Url = query["url"],
                Project = query["project"],
                QueryName = query["queryName"],
                QueryId = query["queryId"],
            };

            foreach (var assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                
            }
        }

        static void Main(Options options)
        {
            var templatePath = Path.GetTempFileName() + ".pbit";
            File.WriteAllBytes(templatePath, Resources.FileTemplate);

            var formula = Resources.QueryTemplate
                .Replace("@URL", options.Url)
                .Replace("@PROJECT", options.Project)
                .Replace("@ID", options.QueryId);

            using (var package = Package.Open(templatePath))
            {
                var mashupPart = package.GetPart(new Uri("/DataMashup", UriKind.Relative));
                var packageComponents = default(PackageComponents);

                using (var stream = mashupPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                {
                    packageComponents = PackageComponents.Deserialize(stream.ReadAllBytes());
                }

                var partStorage = MemoryPackagePartStorage.Deserialize(packageComponents.PartsBytes);
                var permissionsStorage = new MemoryPermissionsStorage(PermissionsSerializer.Deserialize(packageComponents.PermissionBytes));
                var packageMetadataStorage = new MemoryPackageMetadataStorage(true);
                var contentStorate = new MemoryContentStorage();
                var storage = new MemoryPackageStorage(partStorage, permissionsStorage, packageMetadataStorage, contentStorate);
                using (var editor = new PackageEditor(storage))
                {
                    editor.AddFormula(PackageEditor.DefaultSectionName, options.QueryName, formula);

                    editor.Commit();
                    editor.PackageStorage.Commit();

                    packageComponents = new PackageComponents(
                        MemoryPackagePartStorage.Serialize(editor.PackageStorage.Parts),
                        PermissionsSerializer.Serialize(editor.GetPermissions()),
                        PackageMetadataSerializer.Serialize(editor.PackageStorage.Metadata.GetPackageMetadata(), editor.PackageStorage.ContentStorage));
                } 
                using (var stream = mashupPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                {
                    var bytes = packageComponents.Serialize();
                    stream.Write(bytes, 0, bytes.Length);
                }
            }

            Process.Start(templatePath);
        }
    }
}
