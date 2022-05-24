// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");
using System.Reflection;
using System.Xml;
using System;
using SpreadsheetLight;

namespace ConsoleAppLleida // Note: actual namespace depends on the project name.
{
    
    internal class Program
    {
        internal static int contador=0;
        internal static SLDocument osLDocument = new SLDocument();
        internal static System.Data.DataTable dt = new System.Data.DataTable();

        static void Main(string[] args)
        {
            //documento excel
            
            contador++;
            //columnas
            dt.Columns.Add("Id", typeof(string));
            dt.Columns.Add("Fecha", typeof(string));
            dt.Columns.Add("Tipo", typeof(string));
            dt.Columns.Add("Doc_OkKo", typeof(string));
            dt.Columns.Add("Doc_UID", typeof(string));
            dt.Columns.Add("Unidades Certificadas", typeof(string));
            dt.Columns.Add("Dirección Origen", typeof(string));
            dt.Columns.Add("Dirección Destino", typeof(string));
            dt.Columns.Add("Dirección Cc", typeof(string));
            dt.Columns.Add("Estado", typeof(string));
            dt.Columns.Add("Estado Aux", typeof(string));
            dt.Columns.Add("Asunto", typeof(string));
            dt.Columns.Add("Doc_Visualizado", typeof(string));
            dt.Columns.Add("Fecha y hora de visualización", typeof(string));
            dt.Columns.Add("Add_UID", typeof(string));



            string path=Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Resources\prueba.xml");
            //UsingXmlReader("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220524213336");
            //string pathFile = AppDomain.CurrentDomain.BaseDirectory + "miExcel.xlsx";
            string pathFile = AppDomain.CurrentDomain.BaseDirectory + "miExcel.xlsx";
            irExcel(pathFile);
           
        }

        private static void UsingXmlReader(string path)
        {
            contador++;
            string mail_id,mail_date, mail_type, file_doc_model, file_uid, mail_from, mail_to, gstatus, gstatus_aux, mail_subj, add_id;
            XmlReader xmlReader=XmlReader.Create(path);

            while (xmlReader.Read())
            {
                if((xmlReader.NodeType == XmlNodeType.Element)&&(xmlReader.Name== "mail_id"))
                {
                   
                     mail_id = xmlReader.ReadElementContentAsString();
                        Console.WriteLine("mail_id= "+mail_id);
                }
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "mail_date"))
                {
                    mail_date = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("mail_date= " + mail_date);
                                     
                }
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "mail_type"))
                {
                    mail_type = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("mail_type= " + mail_type);
                   
                }
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "file_doc_model"))
                {
                    file_doc_model = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("file_doc_model= " + file_doc_model);

                }
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "file_uid"))
                {
                    file_uid = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("file_uid= " + file_uid);


                }

                //unidades certificadas
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "mail_from"))
                {
                    mail_from = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("mail_from= " + mail_from);


                }
                
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "mail_to"))
                {
                    mail_to = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("mail_subj_dec= " + mail_to);


                }
                //direccion CC
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "gstatus"))
                {
                    gstatus = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("gstatus= " + gstatus);
                   
                }
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "gstatus_aux"))
                {
                    gstatus_aux = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("gstatus_aux= " + gstatus_aux);
                }
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "mail_subj"))
                {
                    mail_subj = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("mail_subj= " + mail_subj);


                }
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "add_id"))
                {
                    add_id = xmlReader.ReadElementContentAsString();
                    Console.WriteLine("add_id= " + add_id);
                    Console.WriteLine("\n");

                }

            }
            
            Console.ReadKey();
        }

        public static void irExcel(string pathFile)
        {
            

            //registros 
            dt.Rows.Add("pepe",19,"hombre");
            dt.Rows.Add("andres", 27, "hombre");
            dt.Rows.Add("Eve", 10, "mujer");

            //donde iniciamos
            osLDocument.ImportDataTable(1,1,dt,true);
            osLDocument.SaveAs(pathFile);

        }

    }
}



