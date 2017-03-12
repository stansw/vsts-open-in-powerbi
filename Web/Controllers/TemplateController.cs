using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using System.Web.Http;
using System.Web.Http.Controllers;
using Microsoft.ApplicationInsights;
using Microsoft.Mashup.Client.Packaging;
using Microsoft.Mashup.Engine.Interface;
using Microsoft.Mashup.Host.Document;
using Microsoft.Mashup.Host.Document.Storage;
using Microsoft.Mashup.Host.Document.Storage.Metadata;
using Microsoft.Mashup.Storage.Memory;
using Web.Properties;

namespace Web.Controllers
{
    public class TemplateController : ApiController
    {
        private TelemetryClient _telemetry;

        protected override void Initialize(HttpControllerContext controllerContext)
        {
            base.Initialize(controllerContext);
            _telemetry = new TelemetryClient();
        }

        // GET _api/Template/flat?url=https://stansw.visualstudio.com&project=vsts-open-in-powerbi&qid=d5349265-9c9d-4808-933a-c3d27b731657&qname=Flat%20Work%20Items
        public HttpResponseMessage Get(string id, string url, string project, string qname, Guid qid)
        {
            // Trace event but do *not* save information which might be considered private.
            _telemetry.TrackEvent("TemplateController.Get", new Dictionary<string, string>()
            {
                ["template"] = id,
                ["url"] = url
            });

            var response = new HttpResponseMessage(HttpStatusCode.OK);
            response.Content = new ByteArrayContent(this.BuildTemplate(id, url, project, qid.ToString("D")));
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = (string.IsNullOrWhiteSpace(qname) ? "Template" : qname) + ".pbit"
            };

            return response;
        }

        public byte[] BuildTemplate(string template, string url, string project, string id)
        {
            var parameters = new Dictionary<string, string>()
            {
                ["url"] = url,
                ["project"] = project,
                ["id"] = id
            };

            var templateBytes = GetTemplateBytes(template);
            var tempFilePath = Path.GetTempFileName();
            File.WriteAllBytes(tempFilePath, templateBytes);

            using (var package = Package.Open(tempFilePath))
            {
                var mashupPart = package.GetPart(new Uri("/DataMashup", UriKind.Relative));
                var packageComponents = default(PackageComponents);

                using (var stream = mashupPart.GetStream(FileMode.Open, FileAccess.ReadWrite))
                {
                    packageComponents = PackageComponents.Deserialize(stream.ReadAllBytes());
                }

                var partStorage = MemoryPackagePartStorage.Deserialize(packageComponents.PartsBytes);
                var permissionsStorage =
                    new MemoryPermissionsStorage(PermissionsSerializer.Deserialize(packageComponents.PermissionBytes));
                var packageMetadataStorage = new MemoryPackageMetadataStorage(true);
                var contentStorate = new MemoryContentStorage();
                var storage = new MemoryPackageStorage(partStorage, permissionsStorage, packageMetadataStorage,
                    contentStorate);
                using (var editor = new PackageEditor(storage))
                {
                    var sourceError = default(IError);
                    var sectionText = default(string);

                    if (!editor.TryGetSectionText(PackageEditor.DefaultSectionName, out sectionText, out sourceError))
                    {
                        throw new InvalidOperationException(sourceError.Message);
                    }

                    foreach (var parameter in parameters)
                    {
                        sectionText = Regex.Replace(sectionText,
                            $@" {Regex.Escape(parameter.Key)}\s*=\s*""[^""]*""",
                            $@" {parameter.Key} = ""{parameter.Value}""");
                    }

                    if (!editor.TrySetSectionText(PackageEditor.DefaultSectionName, sectionText, out sourceError))
                    {
                        throw new InvalidOperationException(sourceError.Message);
                    }

                    editor.Commit();
                    editor.PackageStorage.Commit();

                    packageComponents = new PackageComponents(
                        MemoryPackagePartStorage.Serialize(editor.PackageStorage.Parts),
                        packageComponents.PermissionBytes,
                        packageComponents.MetadataBytes,
                        packageComponents.PermissionBinding);

                    // The following code overrides permissions and metadata. Currently templates depend
                    // on query names and columns returned, thus, this code cannot be used any more.
                    ////packageComponents = new PackageComponents(
                    ////    MemoryPackagePartStorage.Serialize(editor.PackageStorage.Parts),
                    ////    PermissionsSerializer.Serialize(editor.GetPermissions()),
                    ////    PackageMetadataSerializer.Serialize(editor.PackageStorage.Metadata.GetPackageMetadata(), editor.PackageStorage.ContentStorage));
                }
                using (var stream = mashupPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                {
                    var bytes = packageComponents.Serialize();
                    stream.Write(bytes, 0, bytes.Length);
                }
            }

            var result = File.ReadAllBytes(tempFilePath);

            try
            {
                File.Delete(tempFilePath);
            }
            catch (Exception ex)
            {
                _telemetry.TrackException(ex);
            }

            return result;
        }

        private static byte[] GetTemplateBytes(string template)
        {
            switch (template?.ToLower())
            {
                case "flat":
                    return Resources.Flat;
                case "tree":
                    return Resources.Tree;
                case "onehop":
                    return Resources.OneHop;
                default:
                    return Resources.General;
            }
        }

    }
}
