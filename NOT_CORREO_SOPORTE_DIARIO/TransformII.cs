using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


using System.Xml.Serialization;
using System.IO;
using System.Xml.Serialization.Configuration;
//using System.Xml.XmlConfiguration;
using System.Xml;
using System.Xml.Xsl;
using System.Data;

namespace Transform
{
    public class TransformII
    {

        public string IDSOP { get; set; }
        public string Dominio { get; set; }
        public string Grupo { get; set; }
        public string Codigo { get; set; }
        public string Aplicativo { get; set; }
        public string Nombres { get; set; }
        public string FechaCreacion { get; set; }
        public string FechaAsignacion { get; set; }
        public string Antiguedad { get; set; }
        public string Ambiente { get; set; }
        public string Resumen { get; set; }
        public string UsuarioAsignado { get; set; }

        private static List<TransformII> GetData(DataTable dt)
        {
            List<TransformII> listSOP = new List<TransformII>();
            foreach (DataRow dr in dt.Rows)
            {
                listSOP.Add(new TransformII()
                {
                    IDSOP = dr["IDSOP"].ToString(),
                    FechaAsignacion = dr["FechaAsignacion"].ToString(),
                    Antiguedad = dr["Antiguedad"].ToString(),
                    //Dominio = dr["Dominio"].ToString(),
                    Grupo = dr["Grupo"].ToString(),
                    //Codigo = dr["Codigo"].ToString(),
                    Aplicativo = dr["Aplicativo"].ToString(),
                    Ambiente = dr["Ambiente"].ToString(),
                    UsuarioAsignado = dr["UsuarioAsignado"].ToString(),
                    Resumen = dr["Resumen"].ToString(),


                });
            }
            return listSOP;
        }

        public static string TransformXml(DataTable dt, string xslFileName)
        {
            List<TransformII> data = new List<TransformII>();
            data = GetData(dt);

            XmlSerializer xs = new XmlSerializer(typeof(List<TransformII>));


            string xmlString;
            using (StringWriter swr = new StringWriter())
            {
                xs.Serialize(swr, data);
                xmlString = swr.ToString();
            }


            var xd = new XmlDocument();
            xd.LoadXml(xmlString);
            var xslt = new XslCompiledTransform();
            xslt.Load(xslFileName);
            var stm = new MemoryStream();
            xslt.Transform(xd, null, stm);
            stm.Position = 0;
            var sr = new StreamReader(stm);
            return sr.ReadToEnd();

            // return xmlString.ToString();


        }
        public static string CargarLeyenda(string xslFileName)
        {
            var sr = new StreamReader(xslFileName);
            return sr.ReadToEnd();

        }
    }
}
