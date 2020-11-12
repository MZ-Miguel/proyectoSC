using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Data.OleDb;
using System.IO;

using System.IO.Compression;
using System.Text;
using System.Xml;
using io = System.IO;

namespace WebApplication12.Controllers
{
    public class EngineeringRegisterController : Controller
    {
        SqlConnection SCGamma1 = new SqlConnection(ConfigurationManager.ConnectionStrings["SCGamma1"].ConnectionString);
        OleDbConnection Econ;
        // GET: EngineeringRegister
        public ActionResult EngineeringRegister2()
        {
            return View();
        }

        [HttpPost]

        public ActionResult EngineeringRegister2(HttpPostedFileBase file)
        {
            string filename = Guid.NewGuid() + Path.GetExtension(file.FileName);
            string filepath = "/excelfolder/" + filename;
            file.SaveAs(Path.Combine(Server.MapPath("/excelfolder"), filename));
            InsertExceldata(filepath, filename);
            return View();
        }

        //----------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------
        public String LeerExcel()
        {
            StringBuilder sb = new StringBuilder();
            HttpPostedFileBase hpfb = (HttpPostedFileBase)Request.Files["archivo"];
            String nombreArchivo = hpfb.FileName;
            if (io.Path.GetExtension(nombreArchivo).Equals(".xlsx"))
            {
                sb.Append("1»");//Alt+175
                sb.Append(obtenerDatosExcel(hpfb));
            }
            else
            {
                sb.Append("0»");//Alt+175
                sb.Append("No es excel");
            }
            return sb.ToString();
        }

        private string obtenerDatosExcel(HttpPostedFileBase hpfb, Boolean soloPrimeraHoja = false, String sprCelda = "¦", String sprFila = "¬", String sprHoja = "¯")
        {
            String vContenido = "";
            List<String> lHoja = new List<String>();
            List<String> lFila = null;
            List<String> lCelda = null;
            int nRegistros = 0;
            List<string> hojas = new List<string>();
            List<XmlDocument> docs = new List<XmlDocument>();
            List<string> nombres = new List<string>();
            List<string> valores = new List<string>();
            List<string> columnas = new List<string>();
            XmlNodeList xNodos;
            int nNodos = 0;


            using (var zip = new ZipArchive(hpfb.InputStream, ZipArchiveMode.Read))
            {
                XmlDocument doc;
                foreach (var archivoXml in zip.Entries)
                {
                    if (archivoXml.Name == "workbook.xml" || archivoXml.Name == "sharedStrings.xml" || archivoXml.Name.StartsWith("sheet") && io.Path.GetExtension(archivoXml.Name) == ".xml")
                    {
                        using (io.Stream stream = archivoXml.Open())
                        {
                            doc = new XmlDocument();
                            doc.Load(stream);
                            docs.Add(doc);
                            nombres.Add(archivoXml.Name);
                        }
                    }
                }
            }

            if (docs.Count > 0)
            {
                int pos = nombres.IndexOf("sharedStrings.xml");
                if (pos > -1)
                {
                    XmlDocument xdStrings = docs[pos];
                    XmlElement nodoRaizStrings = xdStrings.DocumentElement;
                    XmlNodeList nodosValores = nodoRaizStrings.ChildNodes;
                    if (nodosValores != null)
                    {
                        foreach (XmlNode nodoValor in nodosValores)
                        {
                            if (nodoValor.FirstChild.FirstChild != null)
                            {
                                valores.Add(nodoValor.FirstChild.FirstChild.Value);
                            }
                            else
                            {
                                valores.Add("");
                            }
                        }
                    }
                    pos = nombres.IndexOf("workbook.xml");
                    if (pos > -1)
                    {
                        XmlDocument xdLibro = docs[pos];
                        XmlElement nodoRaizHojas = xdLibro.DocumentElement;
                        XmlNodeList nodosHojas = nodoRaizHojas.GetElementsByTagName("sheet");
                        string id, hoja;
                        if (nodosHojas != null)
                        {
                            int ch = 0;
                            foreach (XmlNode nodoHoja in nodosHojas)
                            {
                                id = nodoHoja.Attributes["r:id"].Value.Replace("rId", "");
                                hoja = nodoHoja.Attributes["name"].Value;
                                pos = nombres.IndexOf("sheet" + id + ".xml");
                                lFila = new List<String>();
                                if (pos > -1)
                                {
                                    XmlDocument xdHoja = docs[pos];
                                    XmlElement nodoRaizHoja = xdHoja.DocumentElement;
                                    XmlNodeList nodosFilas = nodoRaizHoja.GetElementsByTagName("row");
                                    int indice;
                                    string celda, valor;
                                    XmlAttribute tipoString;
                                    int cf = 0;
                                    int cc = 0;
                                    XmlNode nodoCelda = null;
                                    if (nodosFilas != null)
                                    {
                                        foreach (XmlNode nodoFila in nodosFilas)
                                        {
                                            XmlNodeList nodoCeldas = nodoFila.ChildNodes;
                                            if (nodoCeldas != null)
                                            {
                                                if (cf == 0)
                                                {
                                                    columnas = new List<string>();
                                                    nRegistros = nodoCeldas.Count;
                                                    for (int i = 0; i < nRegistros; i++)
                                                    {
                                                        columnas.Add(nodoCeldas[i].Attributes["r"].Value.Replace(nodoFila.Attributes["r"].Value, ""));
                                                    }
                                                }
                                                else
                                                {
                                                    cc = 0;
                                                    lCelda = new List<String>();
                                                    nRegistros = columnas.Count;
                                                    for (int i = 0; i < nRegistros; i++)
                                                    {
                                                        nodoCelda = nodoCeldas[cc];
                                                        if (nodoCelda != null)
                                                        {
                                                            if (columnas[i] == nodoCelda.Attributes["r"].Value.Replace(nodoFila.Attributes["r"].Value, ""))
                                                            {
                                                                celda = nodoCelda.Attributes["r"].Value;
                                                                tipoString = nodoCelda.Attributes["t"];
                                                                valor = "";
                                                                //String
                                                                if (tipoString != null)
                                                                {
                                                                    if (tipoString.Value.Equals("s"))
                                                                    {
                                                                        if (valores != null && valores.Count > 0)
                                                                        {
                                                                            indice = int.Parse(nodoCelda.FirstChild.FirstChild.Value);
                                                                            valor = valores[indice];
                                                                            lCelda.Add(valor);
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        xNodos = nodoCelda.ChildNodes;
                                                                        nNodos = xNodos.Count;
                                                                        //Función Concatenar
                                                                        for (int j = 0; j < nNodos; j++)
                                                                        {
                                                                            if (xNodos[j].Name.Equals("v"))
                                                                            {
                                                                                valor = xNodos[j].FirstChild.Value;
                                                                                lCelda.Add(valor);
                                                                                break;
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                                //Número, Fecha o Función
                                                                else
                                                                {
                                                                    if (nodoCelda.FirstChild != null && nodoCelda.FirstChild.FirstChild != null)
                                                                    {
                                                                        xNodos = nodoCelda.ChildNodes;
                                                                        nNodos = xNodos.Count;
                                                                        //Función o Suma de celdas
                                                                        if (nNodos > 1)
                                                                        {
                                                                            for (int j = 0; j < nNodos; j++)
                                                                            {
                                                                                if (xNodos[j].Name.Equals("v"))
                                                                                {
                                                                                    valor = xNodos[j].FirstChild.Value;
                                                                                    lCelda.Add(valor);
                                                                                    break;
                                                                                }
                                                                            }
                                                                        }
                                                                        //Número o Fecha
                                                                        //else
                                                                        //{
                                                                        //    if (nodoCelda.Attributes["s"] != null)
                                                                        //    {
                                                                        //        foreach (XmlAttribute att in nodoCelda.Attributes)
                                                                        //        {
                                                                        //            //Fecha
                                                                        //            if (att.Value.Equals("3") /*|| att.Value.Equals("4")*/)
                                                                        //            {
                                                                        //                tmpFecha = Double.Parse(nodoCelda.FirstChild.FirstChild.Value);
                                                                        //                oFecha = DateTime.FromOADate(tmpFecha);
                                                                        //                valor = oFecha.ToString("dd/MM/yyyy");
                                                                        //            }
                                                                        //            //Número
                                                                        //            else
                                                                        //            {
                                                                        //                valor = nodoCelda.FirstChild.FirstChild.Value;
                                                                        //            }
                                                                        //        }
                                                                        //        lCelda.Add(valor);
                                                                        //    }
                                                                        //}
                                                                        /*BORRAR SOLO PARA PAMS*/
                                                                        else
                                                                        {
                                                                            //Fecha
                                                                            /*if (nodoCelda.Attributes["s"] != null)
                                                                            {
                                                                                tmpFecha = Double.Parse(nodoCelda.FirstChild.FirstChild.Value);
                                                                                oFecha = DateTime.FromOADate(tmpFecha);
                                                                                valor = oFecha.ToString("dd/MM/yyyy");
                                                                            }*/
                                                                            //Número
                                                                            /*else
                                                                            {*/
                                                                            valor = nodoCelda.FirstChild.FirstChild.Value;
                                                                            //}
                                                                            lCelda.Add(valor);
                                                                        }
                                                                        /**/
                                                                    }
                                                                    else
                                                                    {
                                                                        lCelda.Add(valor);
                                                                    }
                                                                }
                                                                cc++;
                                                            }
                                                            else
                                                            {
                                                                //Celda vacía
                                                                lCelda.Add("");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            //Celda vacía
                                                            lCelda.Add("");
                                                        }
                                                    }
                                                    lFila.Add(String.Join(sprCelda, lCelda));
                                                }
                                            }
                                            cf++;
                                        }
                                    }
                                    lHoja.Add(String.Join(sprFila, lFila));
                                }
                                ch++;
                                if (soloPrimeraHoja) break;
                            }
                        }
                    }
                }
            }
            vContenido = String.Join(sprHoja, lHoja);
            return vContenido;
        }

        //----------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------
        //----------------------------------------------------------------------------------------
        private void ExcelConn(string filepath)
        {
            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", filepath);
            Econ = new OleDbConnection(constr);
        }

        private void InsertExceldata(string fileepath, string filename)

        {
            string fullpath = Server.MapPath("/excelfolder/") + filename;
            ExcelConn(fullpath);
            string query = string.Format("Select * from [{0}]", "PLANEAMITENTO$");
            OleDbCommand Ecom = new OleDbCommand(query, Econ);
            Econ.Open();
            DataSet ds = new DataSet();
            OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);
            Econ.Close();
            oda.Fill(ds);
            DataTable dt = ds.Tables[0];
            SqlBulkCopy objbulk = new SqlBulkCopy(SCGamma1);
            objbulk.DestinationTableName = "ELEMENT";
            objbulk.ColumnMappings.Add("OTC", "OTC");
            objbulk.ColumnMappings.Add("ITEM", "ITEM");
            objbulk.ColumnMappings.Add("CANT", "CANT");
            objbulk.ColumnMappings.Add("MARCA", "MARCA");
            objbulk.ColumnMappings.Add("DESCRIPCION", "DESCRIPCION");
            objbulk.ColumnMappings.Add("LONG", "LONG");
            objbulk.ColumnMappings.Add("AREA", "AREA");
            objbulk.ColumnMappings.Add("PESO", "PESO");
            objbulk.ColumnMappings.Add("SECTOR", "SECTOR");
            objbulk.ColumnMappings.Add("EJE", "EJE");
            objbulk.ColumnMappings.Add("NIVEL", "NIVEL");
            objbulk.ColumnMappings.Add("PISO", "PISO");
            objbulk.ColumnMappings.Add("PRIORIDAD", "PRIORIDAD");
            objbulk.ColumnMappings.Add("BLOQUE", "BLOQUE");
            SCGamma1.Open();
            objbulk.WriteToServer(dt);
            SCGamma1.Close();
        }
    }
}