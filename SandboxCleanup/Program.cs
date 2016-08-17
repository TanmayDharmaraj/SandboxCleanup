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

            //Remove Assembly References
            XmlNode nodes = doc.SelectSingleNode("/a:Solution/a:Assemblies", nsmgr);
            if (nodes != null)
            {
                foreach (XmlNode node in nodes)
                {
                    if (node.Attributes["Location"] == null) continue;
                    string dll = node.Attributes["Location"].Value;

                    if (delete_dll)
                    {
                        //TODO: this code will throw an exception if the dll is set to read-only.
                        File.Delete(Path.Combine(sandbox_folder_path, dll));
                    }

                }
                solution_node.RemoveChild(nodes);
            }

            //Remove Feature ReceiverClass and ReceiverAssembly attributes
            XmlNodeList feature_nodes = doc.SelectNodes("/a:Solution/a:FeatureManifests/a:FeatureManifest", nsmgr);
            if (feature_nodes != null)
            {
                foreach (XmlNode node in feature_nodes)
                {
                    if (node.Attributes["Location"] == null) continue;
                    string feature_location = node.Attributes["Location"].Value;

                    //Get the feature.xml file
                    string feature_path = Path.Combine(sandbox_folder_path, feature_location);
                    XmlDocument feature_xml = new XmlDocument();
                    feature_xml.Load(feature_path);
                    //Load the xml document for feature.xml
                    XmlNodeList f_nodelist = feature_xml.SelectNodes("/a:Feature", nsmgr);
                    
                    
                    foreach (XmlNode f_node in f_nodelist)
                    {
                        //Remove Feature ReceiverClass and ReceiverAssembly attributes
                        XmlAttribute recevierClass_attrib = f_node.Attributes["ReceiverClass"];
                        if (recevierClass_attrib != null)
                            f_node.Attributes.Remove(recevierClass_attrib);

                        XmlAttribute recevierAssembly_attrib = f_node.Attributes["ReceiverAssembly"];
                        if (recevierAssembly_attrib != null)
                            f_node.Attributes.Remove(recevierAssembly_attrib);

                        //Remove ElementFile having .webpart as extention
                        XmlNode element_manifests_parent_node = f_node.SelectSingleNode("/a:Feature/a:ElementManifests", nsmgr);
                        XmlNodeList element_files = f_node.SelectNodes("/a:Feature/a:ElementManifests/a:ElementFile[contains(@Location,'webpart')]", nsmgr);
                        if (element_files != null)
                        {
                            foreach (XmlNode element_file in element_files)
                            {
                                if (element_file.Attributes["Location"] != null)
                                {
                                    string webpart_file_path = Path.Combine(Path.GetDirectoryName(feature_path),element_file.Attributes["Location"].Value);
                                    string webpart_directory_path = Path.GetDirectoryName(webpart_file_path);
                                    Directory.Delete(webpart_directory_path, true);
                                }
                                element_manifests_parent_node.RemoveChild(element_file);
                            }
                        }

                    }

                    

                    //Save feature.xml back to disk
                    feature_xml.Save(feature_path);
                }
            }
            doc.Save(manifest_path);
        }

        private static void DeleteDlls(string path)
        {
            File.Delete(path);
        }

        static void Main(string[] args)
        {

            string root_destination_folder = @"C:\Users\TanmayD\Downloads\Sandbox\Abellio\Extracted";
            string cab_destination_folder = @"C:\Users\TanmayD\Downloads\Sandbox\Abellio\Converted";
            string[] sandbox_solution_paths = new string[] {
                @"C:\Users\TanmayD\Downloads\Sandbox\Abellio\WSP\ACFilters.wsp"
            };
            Directory.CreateDirectory(cab_destination_folder);
            foreach (string sandbox_path in sandbox_solution_paths)
            {
                string solution_name_without_extention = Path.GetFileNameWithoutExtension(sandbox_path);
                string extracted_folder = Path.Combine(root_destination_folder, solution_name_without_extention);

                extract(sandbox_path, extracted_folder);
                CleanManifest(Path.Combine(extracted_folder, "manifest.xml"), delete_dll: true);

                makecab(Path.Combine(cab_destination_folder, Path.GetFileName(sandbox_path)), extracted_folder);
            }
        }
    }
}
