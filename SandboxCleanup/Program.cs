using Microsoft.Deployment.Compression.Cab;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace SandboxCleanup
{
    class Program
    {
        private static void extract(string filepath, string destination_path)
        {
            var cabInfo = new CabInfo(filepath);
            cabInfo.Unpack(destination_path);
        }

        private static void makecab(string destination, string source_folder_path)
        {
            var cab = new CabInfo(destination);
            cab.Pack(source_folder_path, true, Microsoft.Deployment.Compression.CompressionLevel.Normal, null);
        }

        private static void CleanManifest(string manifest_path, bool delete_dll = false)
        {
            string sandbox_folder_path = new FileInfo(manifest_path).Directory.FullName;
            XmlDocument doc = new XmlDocument();
            var nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("a", "http://schemas.microsoft.com/sharepoint/");
            doc.Load(manifest_path);

            XmlNode solution_node = doc.SelectSingleNode("/a:Solution", nsmgr);
            XmlNode nodes = doc.SelectSingleNode("/a:Solution/a:Assemblies", nsmgr);

            foreach (XmlNode node in nodes)
            {
                if (node.Attributes["Location"] == null) continue;
                string dll = node.Attributes["Location"].Value;

                if (delete_dll)
                {
                    File.Delete(Path.Combine(sandbox_folder_path, dll));
                }

            }

            solution_node.RemoveChild(nodes);
            doc.Save(manifest_path);
        }

        private static void DeleteDlls(string path)
        {
            File.Delete(path);
        }

        static void Main(string[] args)
        {
            
            string root_destination_folder = @"C:\Extracted";
            string cab_destination_folder = @"C:\Extracted\CabFiles";
            string[] sandbox_solution_paths = new string[] {
                @"C:\Users\TanmayD\Documents\test\SharePointProject2.wsp"
            };
            Directory.CreateDirectory(cab_destination_folder);
            foreach (string sandbox_path in sandbox_solution_paths)
            {
                string solution_name_without_extention = Path.GetFileNameWithoutExtension(sandbox_path);
                string extracted_folder = Path.Combine(root_destination_folder, solution_name_without_extention);

                extract(sandbox_path, extracted_folder);
                CleanManifest(Path.Combine(extracted_folder, "manifest.xml"), delete_dll: true);
                
                makecab(Path.Combine(cab_destination_folder,Path.GetFileName(sandbox_path)), extracted_folder);
            }
        }
    }
}
