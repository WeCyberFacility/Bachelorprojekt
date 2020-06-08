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

                IEnumerable<XElement> systemarten = GetXElementsByName(xmldoc, "2");
                IEnumerable<XElement> ziele = GetXElementsByName(xmldoc, "7");
                //printeListe("Systemart: ", systemarten);
                getZiele(ziele);
                getSystemarten(systemarten);
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
           
            IEnumerable<XElement> de =
                from el in zielliste.Descendants()
                where el.Name.LocalName == "Text"
                select el;
            foreach (XElement el in de)
                ziele.Add(new Ziel(el.Value));



            Console.WriteLine("Ziele gefunden:");
            
            foreach (Ziel currentziel in ziele)
            {
                Console.WriteLine(currentziel.getName());
            }

            return ziele;
        }

        //Getter für verschiedene Typen in unserem Diagramm:
        private static IEnumerable<Systemart> getSystemarten(IEnumerable<XElement> systemartliste)
        {
            //neue leere Liste
            List<Systemart> systemarts = new List<Systemart>();

            //laufe die unter-Tags durch und suche nach dem Tag "Text"
            IEnumerable<XElement> de =
                from el in systemartliste.Descendants()
                where el.Name.LocalName == "Text"
                select el;
            //für jedes Tag XElement erstelle eine neue Systemart und füge es der Systemartsliste hinzu
            foreach (XElement el in de)
            {
                systemarts.Add(new Systemart(el.Value));
            }


            //Durchlaufe nun wieder die gefundenen Systemart shapes und nehme jeweils die IDs aus ihnen raus uznd füge es der zugehörigen Systemart in der Systemartliste hinzu
            int counter = 0;
            foreach (XElement systemartxml in systemartliste)
            {
                Console.WriteLine(systemartxml.Attribute("ID").Value);
                systemarts.ElementAt(counter).setId(systemartxml.Attribute("ID").Value);
                
                counter++;
            }


            //Gib alle Systemarten aus:
            Console.WriteLine("Systemarten gefunden:");

            foreach (Systemart currentziel in systemarts)
            {
                Console.WriteLine(currentziel.getName() + " ID: " + currentziel.getId());
            }

            return systemarts;
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

    class Systemart
    {
        string name;
        string id;
        XElement xElement;

        public Systemart(string name)
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
        public string getId()
        {
            return this.id;
        }
        public void setId(string newid)
        {
            this.id = newid;
        }
    }

}
