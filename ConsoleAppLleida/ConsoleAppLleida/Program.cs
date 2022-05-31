// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World!");
using System.Reflection;
using System.Xml;
using System;
using SpreadsheetLight;
using System.Windows;
using Microsoft.Win32;

namespace ConsoleAppLleida // Note: actual namespace depends on the project name.
{
    
    internal class Program
    {
        internal static int contador=0,marce=0;
        internal static SLDocument osLDocument = new SLDocument();
        internal static System.Data.DataTable dt = new System.Data.DataTable();
        internal static string mail_id, mail_date, mail_type, file_doc_model, file_uid, unidades_certificadas, mail_from, mail_to,direccion_CC, gstatus, gstatus_aux, mail_subj, add_id, add_displaydate, add_uid;


        static void Main(string[] args)
        {
            //documento excel
            DateTime fechaActual = DateTime.Today;
            //Console.WriteLine(fechaActual.Year);
            int ano2=2021,mes2=0,dia2;
            bool condicion=false;
            string ano, mes, dia;
           /* do
            {
                Console.WriteLine("Por favor ingresar fecha inicial desde donde desea el reporte.");
                do
                {
                    
                    Console.WriteLine("Ingrese año: ");
                    ano = Console.ReadLine();
                    try
                    {
                        ano2 = Int32.Parse(ano);
                        if (2021 <= ano2 && ano2 <= fechaActual.Year)
                        {
                            Console.WriteLine("correcto");
                            condicion = true;
                        }
                        else
                        {
                            Console.Clear();
                            Console.WriteLine("Ingrese un número entre 2021 y " + fechaActual.Year.ToString());
                        }
                    }
                    catch (Exception e)
                    {
                        Console.Clear();
                        Console.WriteLine("Ingrese un número entre 2021 y "+fechaActual.Year.ToString());
                    }

                } while (condicion == false);
                //mes
                condicion = false;
                do
                {
                    Console.WriteLine("Ingrese mes: ");
                    mes = Console.ReadLine();
                    try
                    {
                        mes2 = Int32.Parse(mes);

                        if (ano2==fechaActual.Year)
                        {
                            if (1 <= mes2 && mes2 <= fechaActual.Month)
                            {

                                condicion = true;
                            }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine("El mes ingresado es mayor al actual");
                            }
                        }
                        else
                        {
                            if (1 <= mes2 && mes2 <= 12)
                            {

                                condicion = true;
                            }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine("Ingrese un número entre 1 y 12");
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.Clear();
                        Console.WriteLine("Ingrese un número entre 1 y 12");
                    }

                } while (condicion == false);

                //dia
                /*condicion = false;
                do
                {
                    Console.WriteLine("Ingrese dia: ");
                    String dia;
                    dia = Console.ReadLine();
                    try
                    {
                        dia2 = Int32.Parse(dia);
                        if (1 <= dia2 && dia2 <= 31)
                        {
                            Console.WriteLine("correcto");
                            condicion = true;
                        }
                        else
                        {
                            Console.Clear();
                            Console.WriteLine("Ingrese un número entre 1 y 12");
                        }
                    }
                    catch (Exception e)
                    {
                        Console.Clear();
                        Console.WriteLine("Ingrese un número entre 1 y 12");
                    }

                } while (condicion == false);

                

            } while (condicion == false);*/

            
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


            //string hola = "0123456789abcd";
            //Console.WriteLine(hola);
            //hola = hola.Substring(6,2) + "/" + hola.Substring(4,2) + "/" + hola.Substring(0, 4) + " " + hola.Substring(8,2) + ":" + hola.Substring(10,2) + ":" + hola.Substring(12,2);
            //hola = hola.Substring(5,2);//+"/"+hola.Substring(0,4);
            //Console.WriteLine(hola);

            //string path=Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Resources\prueba.xml");
            int mesFin = 0, anoFin = 0;
            string fechaInicio = "",fechaFin=""; 
            /*if (mes2<10)
            {
                mes = "0" + mes;
                fechaInicio= ano + mes + "01070000";
                               
            }           
            else 
            {

                fechaInicio = ano + mes + "01070000";
            }
            
            mesFin = mes2 + 1;
            if (mesFin < 10)
            {
                fechaFin = ano + "0" + mesFin.ToString() + "01070000";
                
            }
            else if (mesFin == 13)
            {
                anoFin=ano2 + 1;
                mesFin = 1;
                fechaFin = anoFin + "0" + mesFin.ToString() + "01070000";
            }
            else
            {
                fechaFin = ano + mesFin.ToString() + "01070000";
            }
            Console.WriteLine(fechaInicio);
            Console.WriteLine("fechafin: "+fechaFin);*/
            Console.WriteLine("Espere...");
            //string pathFecha = "https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min="+fechaInicio+"&mail_date_max="+fechaFin;
           
            string pathFecha = "https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220101070000&mail_date_max=20220201070000";
            /*UsingXmlReader(pathFecha);
            contador = 0;
            Console.WriteLine("\nEspere...");
            pathFecha = "https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220201070000&mail_date_max=20220301070000";
            UsingXmlReader(pathFecha);
            contador = 0;
            Console.WriteLine("\nEspere...");
            pathFecha = "https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220301070000&mail_date_max=20220401070000";
            UsingXmlReader(pathFecha);
            contador = 0;
            Console.WriteLine("\nEspere...");
            pathFecha = "https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220401070000&mail_date_max=20220501070000";
            UsingXmlReader(pathFecha);
            contador = 0;*/
            Console.WriteLine("\nEspere...");
            pathFecha = "https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220501070000&mail_date_max=20220601070000";
            UsingXmlReader(pathFecha);
            contador = 0;
            Console.WriteLine("\nEspere...");
            pathFecha = "https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_date_min=20220601070000";
            UsingXmlReader(pathFecha);
                      
            //porNodos("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_id=83626454");
            string pathFile = AppDomain.CurrentDomain.BaseDirectory + "todin.xlsx";
            pathFile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+"\\LleidaEnero_Junio.xlsx";
            
            //string pathFile = Environment + "holaa.xlsx";
            //string pathFile = AppDomain.CurrentDomain.BaseDirectory + "miExcel.xlsx";
            irExcel(pathFile);
            Console.WriteLine("El archivo se guardó en: " + pathFile);

        }

        private static void UsingXmlReader(string path)
        {
            
            XmlReader xmlReader=XmlReader.Create(path);

            while (xmlReader.Read())
            {

                if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "mail_id"))
                {
                    contador++;
                }
                if(contador == 2)
                {
                    
                    dt.Rows.Add(mail_id, mail_date, mail_type, file_doc_model, file_uid, unidades_certificadas, mail_from, mail_to, direccion_CC, gstatus, gstatus_aux, mail_subj, add_id, add_displaydate,add_uid);
                    add_displaydate = "";
                    add_uid = "";
                    add_id = "";
                    Console.WriteLine("envio datos");
                    contador=contador-1;
                }

                    if ((xmlReader.NodeType == XmlNodeType.Element)&&(xmlReader.Name== "mail_id"))
                {
                   
                     mail_id = xmlReader.ReadElementContentAsString();
                        Console.WriteLine("mail_id= "+mail_id);
                }
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "mail_date"))
                {
                    mail_date = xmlReader.ReadElementContentAsString();
                    mail_date = mail_date.Substring(6, 2) + "/" + mail_date.Substring(4, 2) + "/" + mail_date.Substring(0, 4) + " " + mail_date.Substring(8, 2) + ":" + mail_date.Substring(10, 2) + ":" + mail_date.Substring(12, 2);
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
                    unidades_certificadas = "1.00";
                    direccion_CC = "correo@certificado.lleida.net";


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
                    if (add_id!="")
                    {
                        add_uid = "E" + add_id + "-R";
                        add_id = "Displayed";
                        
                    }
                    else
                    {
                        add_id = "";
                    }
                    Console.WriteLine("add_id= " + add_id);
                    

                }
                else if ((xmlReader.NodeType == XmlNodeType.Element) && (xmlReader.Name == "add_displaydate"))
                {
                    add_displaydate = xmlReader.ReadElementContentAsString();
                    add_displaydate = add_displaydate.Substring(6, 2) + "/" + add_displaydate.Substring(4, 2) + "/" + add_displaydate.Substring(0, 4) + " " + add_displaydate.Substring(8, 2) + ":" + add_displaydate.Substring(10, 2) + ":" + add_displaydate.Substring(12, 2);
                    Console.WriteLine("add_displaydate= " + add_id);
                    Console.WriteLine("\n");

                }


            }
            dt.Rows.Add(mail_id, mail_date, mail_type, file_doc_model, file_uid, unidades_certificadas, mail_from, mail_to, direccion_CC, gstatus, gstatus_aux, mail_subj, add_id, add_displaydate, add_uid);

        }


        //con xmlDocument++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        /*private static void porNodos(string path)
        {
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load("https://tsa.lleida.net/cgi-bin/mailcertapi.cgi?action=list_pdf&user=sodigsa@ec&password=TIiANcmymJ&mail_id=83626454");

            foreach (XmlNode xmlNode in xmlDocument.DocumentElement.ChildNodes[0].ChildNodes)
            {
                Console.WriteLine(xmlNode.GetNamespaceOfPrefix);
            }
        }*/


        public static void irExcel(string pathFile)
        {


            //registros 
            /* dt.Rows.Add("pepe",19,"hombre");
             dt.Rows.Add("andres", 27, "hombre");
             dt.Rows.Add("Eve", 10, "mujer");*/

            //donde iniciamos
            Console.WriteLine("Espere XD...");
            osLDocument.ImportDataTable(1,1,dt,true);
            osLDocument.SaveAs(pathFile);
            Console.WriteLine("presiona algo");
            Console.ReadKey();

        }

    }
}



