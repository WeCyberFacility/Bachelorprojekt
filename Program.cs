using System;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.IO.Packaging;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using System.Collections;

namespace ConsoleApp1
{
    class Program
    {

        //Legende zu IDs in Shapes:
        /* 2: Systemart
         * 5: Actor in einem Zielmodell
         * 7: Goals
         * 8: Softgoals
         * 10: Decompositions (Zerteilung der Ziele in Unterziele)
         * 12: Einflüsse auf Softgoals
         * 14: Excludes oder Requires
         * */


        static void Main(string[] args)
        {
            Console.WriteLine("Bitte drück eine Taste um den Scan deiner vsdx-Datei zu starten!");
            Console.ReadKey();

            // Open the Visio file in a Package object.
            using (Package visioPackage = OpenPackage("hio.vsdx",
                Environment.SpecialFolder.Desktop))
            {
                PackagePart documentPart = getPage1(visioPackage);
                XDocument xmldoc = GetXMLFromPart(documentPart);
                //Console.WriteLine(xmldoc); printed das gesamte XML Dokument aus!

                //Console.WriteLine("Systemarten im Dokument:");
                IEnumerable<XElement> systemarten = GetXElementsByName(xmldoc, "2");
                printeListe("Systemarten", systemarten);

                //Console.WriteLine("Ziele im Dokument:");
                IEnumerable<XElement> ziele = GetXElementsByName(xmldoc, "7");
                printeListe("Ziele", ziele);
                //getZiele(ziele);
            }


        }


        private static Package OpenPackage(string fileName,
    Environment.SpecialFolder folder)
        {
            Package visioPackage = null;
            // Get a reference to the location 
            // where the Visio file is stored.
            string directoryPath = System.Environment.GetFolderPath(
                folder);
            DirectoryInfo dirInfo = new DirectoryInfo(directoryPath);
            // Get the Visio file from the location.
            FileInfo[] fileInfos = dirInfo.GetFiles(fileName);
            if (fileInfos.Count() > 0)
            {
                FileInfo fileInfo = fileInfos[0];
                string filePathName = fileInfo.FullName;
                // Open the Visio file as a package with
                // read/write file access.
                visioPackage = Package.Open(
                    filePathName,
                    FileMode.Open,
                    FileAccess.ReadWrite);
            }
            // Return the Visio file as a package.
            return visioPackage;
        }


        private static PackagePart getPage1(Package filePackage)
        {

            Uri uriDocumentTarget = new Uri("/visio/pages/page1.xml", UriKind.Relative);
            Uri page1uri = PackUriHelper.ResolvePartUri(
            new Uri("/", UriKind.Relative), uriDocumentTarget);

            PackagePart shape1package = filePackage.GetPart(partUri: page1uri);
            Console.WriteLine("Package part URI: {0}", shape1package.Uri);
            Console.WriteLine("Content type: {0}", shape1package.ContentType.ToString());
            return shape1package;

        }

        private static XDocument GetXMLFromPart(PackagePart packagePart)
        {
            XDocument partXml = null;
            // Open the packagePart as a stream and then 
            // open the stream in an XDocument object.
            Stream partStream = packagePart.GetStream();
            partXml = XDocument.Load(partStream);
            return partXml;
        }

        private static IEnumerable<XElement> GetXElementsByName(
    XDocument packagePart, String id)
        {
            // Construct a LINQ query that selects elements by their element type.
            IEnumerable<XElement> elements =
                from element in packagePart.Descendants()
                where element.Name.LocalName == "Shape" && element.Attribute("Master").Value == id
                select element;
            // Return the selected elements to the calling code.
            return elements.DefaultIfEmpty(null);
        }


        private static void printeListe(String art, IEnumerable<XElement> liste)
        {
            Console.WriteLine(art + ":");
            foreach (XElement x in liste)
            {
                Console.WriteLine(x);
            }
        }


        private static IEnumerable<Ziel> getZiele(IEnumerable<XElement> zielliste)
        {
            List<Ziel> ziele = new List<Ziel>();
            foreach (XElement x in zielliste)
            {
                ziele.Add(new Ziel(x.Element("Text").Value));
            }

            Console.WriteLine("Ziele gefunden:");
            Console.WriteLine(ziele);
            return ziele;
        }

    }




    class Ziel
    {
        string name;
 
        public Ziel(string name)
        {
            this.name = name;
        }

        public string getName()
        {
            return this.name;
        }
        public void setName(string newname)
        {
            this.name = newname;
        }
    }

}
