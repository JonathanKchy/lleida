// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");
using System.Reflection;
using System.Xml;
using System;

namespace ConsoleAppLleida // Note: actual namespace depends on the project name.
{
    
    internal class Program
    {
       
        static void Main(string[] args)
        {
            
            string path=Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Resources\prueba.xml");
            UsingXmlReader("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220524193824");
        }

        private static void UsingXmlReader(string path)
        {
             int contador=0;
        XmlReader xmlReader=XmlReader.Create(path);

            while (xmlReader.Read())
            {
                if((xmlReader.NodeType == XmlNodeType.Element)&&(xmlReader.Name== "mail_id"))
                {
                    if (!xmlReader.HasAttributes)
                    {
                        contador++;
                        Console.WriteLine(contador.ToString());
                        Console.WriteLine(xmlReader.ReadElementContentAsString());
                    }
                }
            }
            Console.ReadKey();
        }
    }
}



