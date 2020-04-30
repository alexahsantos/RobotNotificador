using log4net;
using log4net.Config;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NOT_CORREO_SOPORTE_DIARIO
{
    static class Program
    {
        private static readonly ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);


        public static string cs_Remedy = ConfigurationManager.ConnectionStrings["csTCS_AM"].ConnectionString;
        //ConfigurationSettings.AppSettings[""]; //"Data Source=PCONSSQLD04;Initial Catalog=PruebaRemedy;Integrated Security=SSPI;User ID=credito.bcp.com.pe/USERREMEDY;Password=Dene@2018!";


        //public static string plantilla_SPVDCA = "";
        public static string PlantillaSoporteVerde="";
        public static string PlantillaSoporteNaranja="";
        public static string PlantillaSoporteRojo="";
        public static string PlantillaSoporte = "";



        public static string IDSOPSPV_DCA = "";

        public static DataSet listadoCorreosTo = new DataSet();
        public static DataSet listadoCorreosCC = new DataSet();

        public static DataSet listadoCorreos = new DataSet();
        public static DataSet listadoSPV_DCA = new DataSet();

       
        public static DataSet listadoSopInc = new DataSet();
        public static DataSet listadoSopIncVerde = new DataSet();
        public static DataSet listadoSopIncNaranja = new DataSet();
        public static DataSet listadoSopIncRojo = new DataSet();
        public static DataSet listadoSopWos = new DataSet();
        public static DataSet listadoSopWosVerde = new DataSet();
        public static DataSet listadoSopWosNaranja = new DataSet();
        public static DataSet listadoSopWosRojo = new DataSet();

        public static string CorreoTo = "";
        public static string CorreoCC = "";



        public static int cantidad = -1;
        //public static string strBodySPVDCA = "";
        public static string strBodySopInc = "";
        public static string strBodySopIncVerde = "";
        public static string strBodySopIncNaranja = "";
        public static string strBodySopIncRojo = "";
        public static string strBodySopWos = "";
        public static string strBodySopWosVerde = "";
        public static string strBodySopWosNaranja = "";
        public static string strBodySopWosRojo = "";

        public static string email = "";// System.Configuration.ConfigurationManager.AppSettings["email"].ToString();// System.Configuration.ConfigurationSettings.AppSettings["email"].ToString();
        public static string password = "";// System.Configuration.ConfigurationManager.AppSettings["password"].ToString();//System.Configuration.ConfigurationSettings.AppSettings["password"].ToString();
        public static string matricula = "";//System.Configuration.ConfigurationManager.AppSettings["matricula"].ToString(); //System.Configuration.ConfigurationSettings.AppSettings["matricula"].ToString();




        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            XmlConfigurator.Configure();
            _log.Info(new { Body = "Inició Programa" });

            try
            {

                using (SqlConnection cn = new SqlConnection(Program.cs_Remedy))
                {
                    cn.Open();

                    int hora = DateTime.Now.Hour;
                    //Notificación por Colores
                    if (hora == 8)
                       {
                        EnviarEmailSoporteNot_II();
                     }

                    if (hora == 9)
                    {
                        EnviarEmailSoporteNot_I();
                    }
                 
                    cn.Close();
                }
                _log.Info(new { Body = "Finalizó Programa" });
            }
            catch (Exception e)
            {
                /*
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
                */
                _log.Error(new { Body = "" + e.Message.ToString() });
                _log.Info(new { Body = "Finalizó Programa" });

            }

        }


        private static string GetCorreosML(DataTable dt)
        {
            string correos = "";
            foreach (DataRow dr in dt.Rows)
            {
                correos += dr["Correo"].ToString() + ",";
            }
            return correos;
        }
        private static string GetCorreosAnalista(DataTable dt)
        {
            string correos = "";
            foreach (DataRow dr in dt.Rows)
            {
                correos += dr["Correo"].ToString() + ",";
            }
            return correos;
        }
        private static string GetCorreosSupervisor(DataTable dt)
        {
            string correos = "";
            foreach (DataRow dr in dt.Rows)
            {
                correos += dr["Correo"].ToString() + ",";
            }
            return correos;
        }

        public static void EnviarEmailSoporteNot_I()
        {


            using (SqlConnection cn = new SqlConnection(Program.cs_Remedy))
            {


                cn.Open();


                PlantillaSoporte = new SqlCommand("SELECT ValorParametro FROM Parametros P WHERE NombreParametro = 'PLANTILLA_NOT_SOPORTE'", cn).ExecuteScalar().ToString();

                email = new SqlCommand("select ValorParametro from Parametros where NombreParametro='CORREO_EMISOR_NOTIFICACIONES' ", cn).ExecuteScalar().ToString();
                password = new SqlCommand("SELECT ValorParametro FROM Parametros WHERE NombreParametro = 'PWDREMEDY1' ", cn).ExecuteScalar().ToString();
                matricula = new SqlCommand("SELECT ValorParametro FROM Parametros WHERE NombreParametro = 'USERREMEDY1' ", cn).ExecuteScalar().ToString();

                //or MATRICULA = 'S20872' or MATRICULA = 'S20860'  or MATRICULA = 'S34900'//WHERE MATRICULA = 'S75946' or MATRICULA = 'S73659'  
                listadoCorreos = new DataSet();

                listadoCorreosTo = new DataSet();
                listadoCorreosCC = new DataSet();

                // SqlCommand cmdCorreos = new SqlCommand("EXEC ENVIO_CORREO_US 2", cn);

                SqlCommand cmdCorreosTo = new SqlCommand("EXEC ENVIO_SOPORTE_NOT 5", cn);
                SqlCommand cmdCorreosCC = new SqlCommand("EXEC ENVIO_SOPORTE_NOT 6", cn);
                //SqlCommand cmdCorreos = new SqlCommand("select correo from Usuario  where  Codigo='S73659'  ", cn);


                //************************************************************************//
                SqlDataAdapter adpCorreosTo = new SqlDataAdapter(cmdCorreosTo);
                adpCorreosTo.Fill(listadoCorreosTo, "myTable");

                SqlDataAdapter adpCorreosCC = new SqlDataAdapter(cmdCorreosCC);
                adpCorreosCC.Fill(listadoCorreosCC, "myTable");
                //************************************************************************//
                //SqlDataAdapter adpCorreos = new SqlDataAdapter(cmdCorreos);
                //adpCorreos.Fill(listadoCorreos, "myTable");



                listadoSopInc = new DataSet();
                listadoSopWos = new DataSet();

                strBodySopInc = "";
                strBodySopWos = "";


                /************Incidentes ************/
                SqlCommand cmdListaSopInc = new SqlCommand("ENVIO_SOPORTE_NOT", cn);
                cmdListaSopInc.CommandType = CommandType.StoredProcedure;
                cmdListaSopInc.Parameters.Add("@parametro", SqlDbType.Int).Value = 1;
                cmdListaSopInc.ExecuteNonQuery();

                SqlDataAdapter adpListadoSopInc = new SqlDataAdapter(cmdListaSopInc);
                adpListadoSopInc.Fill(listadoSopInc, "myTable");
                /************************************************/


                /************Wos************/
                SqlCommand cmdListaSopWos = new SqlCommand("ENVIO_SOPORTE_NOT", cn);
                cmdListaSopWos.CommandType = CommandType.StoredProcedure;
                cmdListaSopWos.Parameters.Add("@parametro", SqlDbType.Int).Value = 2;
                cmdListaSopWos.ExecuteNonQuery();

                SqlDataAdapter adpListadoSopWos = new SqlDataAdapter(cmdListaSopWos);
                adpListadoSopWos.Fill(listadoSopWos, "myTable");
                /************************************************/
                cn.Close();

                string correosAnalista = GetCorreosML(listadoCorreosTo.Tables[0]);
                string correosSupervisor = GetCorreosML(listadoCorreosCC.Tables[0]);

                if (listadoSopInc.Tables[0].Rows.Count > 0)
                {

                    strBodySopInc = Transform.TransformSoporte.TransformXml(listadoSopInc.Tables[0], PlantillaSoporte.ToString()).ToString();
                }

                if (listadoSopWos.Tables[0].Rows.Count > 0)
                {

                    strBodySopWos = Transform.TransformSoporte.TransformXml(listadoSopWos.Tables[0], PlantillaSoporte.ToString()).ToString();
                }


                if (listadoSopInc.Tables[0].Rows.Count <= 0 && listadoSopInc.Tables[0].Rows.Count <= 0)
                {
                    return;
                }

                string ToMail = correosAnalista;
                if (ToMail.Replace(",", "").Trim().Length == 0)
                    ToMail = correosSupervisor;


                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                //service.Credentials = new WebCredentials(Dts.Variables["Matricula"].Value.ToString(), Dts.Variables["Password"].Value.ToString(), "BCPDOM");
                //service.Credentials = new WebCredentials(matricula, password, "BCPDOM");
                service.TraceEnabled = true;
                service.TraceFlags = TraceFlags.All;
                service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                service.Credentials = new WebCredentials(email, password);
                //service.AutodiscoverUrl(email);
                //service.AutodiscoverUrl("eboluarte@bcp.com.pe", RedirectionUrlValidationCallback);
                EmailMessage oMail = new EmailMessage(service);
                oMail.Subject = "[Notificación Automática] - Lista de INC/WO Activos";


                string bodyCompleto = "Estimados, </br></br>";
                bodyCompleto += "Detalle de Incidentes:</br></br>" + strBodySopInc + "</br></br>Detalle de Ordenes de Trabajos:</br></br>"
                    + strBodySopWos
                    + "</br></br>Quedamos pendiente ante cualquier consulta o comentario.</br>Gracias.</br>";

                // bodyCompleto += "</br></br><strong><span >RECORDAR:</span> En los nuevos indicadores las estimaciones DEBEN cerrarse en 6 días.<strong>";
                bodyCompleto += "</br>Saludos.</br>";

                oMail.Body = new MessageBody(bodyCompleto.ToString());
                oMail.Importance = Importance.High;


                string[] correos;
                correos = correosAnalista.Split(',');
                foreach (string correo in correos)
                    if (correo.Length > 0)
                        oMail.ToRecipients.Add(correo);


                string[] correosCc;
                correosCc = correosSupervisor.Split(',');
                foreach (string correoCc in correosCc)
                    if (correoCc.Length > 0)
                        oMail.CcRecipients.Add(correoCc);

                oMail.BccRecipients.Add(email);
                oMail.Body.BodyType = BodyType.HTML;

                _log.Info(new { Body = "Enviando Correo de Correo de Notificación Soporte Diario Banco" });

                oMail.Send();
                _log.Info(new { Body = "Finalizó Envio de  Correo de Notificación Soporte Diario Banco" });



            }

        }
        public static void EnviarEmailSoporteNot_II()
        {


            using (SqlConnection cn = new SqlConnection(Program.cs_Remedy))
            {


                cn.Open();
                PlantillaSoporteVerde = new SqlCommand("SELECT ValorParametro FROM Parametros P WHERE NombreParametro = 'PLANTILLA_NOT_SOPORTE_VERDE'", cn).ExecuteScalar().ToString();
                PlantillaSoporteNaranja = new SqlCommand("SELECT ValorParametro FROM Parametros P WHERE NombreParametro = 'PLANTILLA_NOT_SOPORTE_NARANJA'", cn).ExecuteScalar().ToString();
                PlantillaSoporteRojo = new SqlCommand("SELECT ValorParametro FROM Parametros P WHERE NombreParametro = 'PLANTILLA_NOT_SOPORTE_ROJO'", cn).ExecuteScalar().ToString();

                email = new SqlCommand("select ValorParametro from Parametros where NombreParametro='CORREO_EMISOR_NOTIFICACIONES' ", cn).ExecuteScalar().ToString();
                password = new SqlCommand("SELECT ValorParametro FROM Parametros WHERE NombreParametro = 'PWDREMEDY1' ", cn).ExecuteScalar().ToString();
                matricula = new SqlCommand("SELECT ValorParametro FROM Parametros WHERE NombreParametro = 'USERREMEDY1' ", cn).ExecuteScalar().ToString();


                //or MATRICULA = 'S20872' or MATRICULA = 'S20860'  or MATRICULA = 'S34900'//WHERE MATRICULA = 'S75946' or MATRICULA = 'S73659'  
                listadoCorreos = new DataSet();

                listadoCorreosTo = new DataSet();
                listadoCorreosCC = new DataSet();

                // SqlCommand cmdCorreos = new SqlCommand("EXEC ENVIO_CORREO_US 2", cn);

                SqlCommand cmdCorreosTo = new SqlCommand("EXEC ENVIO_SOPORTE_NOT 7", cn);
                SqlCommand cmdCorreosCC = new SqlCommand("EXEC ENVIO_SOPORTE_NOT 8", cn);
                //SqlCommand cmdCorreos = new SqlCommand("select correo from Usuario  where  Codigo='S73659'  ", cn);


                //************************************************************************//
                SqlDataAdapter adpCorreosTo = new SqlDataAdapter(cmdCorreosTo);
                adpCorreosTo.Fill(listadoCorreosTo, "myTable");

                SqlDataAdapter adpCorreosCC = new SqlDataAdapter(cmdCorreosCC);
                adpCorreosCC.Fill(listadoCorreosCC, "myTable");
                //************************************************************************//
                //SqlDataAdapter adpCorreos = new SqlDataAdapter(cmdCorreos);
                //adpCorreos.Fill(listadoCorreos, "myTable");



                listadoSopIncVerde = new DataSet();
                listadoSopIncNaranja = new DataSet();
                listadoSopIncRojo = new DataSet();
                listadoSopWosVerde = new DataSet();
                listadoSopWosNaranja = new DataSet();
                listadoSopWosRojo = new DataSet();

                strBodySopIncVerde = "";
                strBodySopIncNaranja = "";
                strBodySopIncRojo = "";
                strBodySopWosVerde = "";
                strBodySopWosNaranja = "";
                strBodySopWosRojo = "";


                /************Incidentes ************/
                SqlCommand cmdListaSopIncVerde = new SqlCommand("ENVIO_SOPORTE_NOT", cn);
                cmdListaSopIncVerde.CommandType = CommandType.StoredProcedure;
                cmdListaSopIncVerde.Parameters.Add("@parametro", SqlDbType.Int).Value = 10;
                cmdListaSopIncVerde.ExecuteNonQuery();

                SqlDataAdapter adpListadoSopIncVerde = new SqlDataAdapter(cmdListaSopIncVerde);
                adpListadoSopIncVerde.Fill(listadoSopIncVerde, "myTable");


                SqlCommand cmdListaSopIncNaranja = new SqlCommand("ENVIO_SOPORTE_NOT", cn);
                cmdListaSopIncNaranja.CommandType = CommandType.StoredProcedure;
                cmdListaSopIncNaranja.Parameters.Add("@parametro", SqlDbType.Int).Value = 11;
                cmdListaSopIncNaranja.ExecuteNonQuery();

                SqlDataAdapter adpListadoSopIncNaranja = new SqlDataAdapter(cmdListaSopIncNaranja);
                adpListadoSopIncNaranja.Fill(listadoSopIncNaranja, "myTable");

                SqlCommand cmdListaSopIncRojo = new SqlCommand("ENVIO_SOPORTE_NOT", cn);
                cmdListaSopIncRojo.CommandType = CommandType.StoredProcedure;
                cmdListaSopIncRojo.Parameters.Add("@parametro", SqlDbType.Int).Value = 12;
                cmdListaSopIncRojo.ExecuteNonQuery();

                SqlDataAdapter adpListadoSopIncRojo = new SqlDataAdapter(cmdListaSopIncRojo);
                adpListadoSopIncRojo.Fill(listadoSopIncRojo, "myTable");

                /************************************************/


                /************Wos************/
                SqlCommand cmdListaSopWosVerde = new SqlCommand("ENVIO_SOPORTE_NOT", cn);
                cmdListaSopWosVerde.CommandType = CommandType.StoredProcedure;
                cmdListaSopWosVerde.Parameters.Add("@parametro", SqlDbType.Int).Value = 13;
                cmdListaSopWosVerde.ExecuteNonQuery();

                SqlDataAdapter adpListadoSopWosVerde = new SqlDataAdapter(cmdListaSopWosVerde);
                adpListadoSopWosVerde.Fill(listadoSopWosVerde, "myTable");

                SqlCommand cmdListaWosNaranja = new SqlCommand("ENVIO_SOPORTE_NOT", cn);
                cmdListaWosNaranja.CommandType = CommandType.StoredProcedure;
                cmdListaWosNaranja.Parameters.Add("@parametro", SqlDbType.Int).Value = 14;
                cmdListaWosNaranja.ExecuteNonQuery();

                SqlDataAdapter adpListadoWosNaranja = new SqlDataAdapter(cmdListaWosNaranja);
                adpListadoWosNaranja.Fill(listadoSopWosNaranja, "myTable");

                SqlCommand cmdListaSopWosRojo = new SqlCommand("ENVIO_SOPORTE_NOT", cn);
                cmdListaSopWosRojo.CommandType = CommandType.StoredProcedure;
                cmdListaSopWosRojo.Parameters.Add("@parametro", SqlDbType.Int).Value = 15;
                cmdListaSopWosRojo.ExecuteNonQuery();

                SqlDataAdapter adpListadoWosIncRojo = new SqlDataAdapter(cmdListaSopWosRojo);
                adpListadoWosIncRojo.Fill(listadoSopWosRojo, "myTable");
                /************************************************/

                cn.Close();

                string correosAnalista = GetCorreosML(listadoCorreosTo.Tables[0]);
                string correosSupervisor = GetCorreosML(listadoCorreosCC.Tables[0]);

                /*INCIDENTES*/

                if (listadoSopIncVerde.Tables[0].Rows.Count > 0)
                {

                    strBodySopIncVerde = Transform.TransformII.TransformXml(listadoSopIncVerde.Tables[0], PlantillaSoporteVerde.ToString()).ToString();
                }


                if (listadoSopIncNaranja.Tables[0].Rows.Count > 0)
                {

                    strBodySopIncNaranja = Transform.TransformII.TransformXml(listadoSopIncNaranja.Tables[0], PlantillaSoporteNaranja.ToString()).ToString();
                }
                if (listadoSopIncRojo.Tables[0].Rows.Count > 0)
                {

                    strBodySopIncRojo = Transform.TransformII.TransformXml(listadoSopIncRojo.Tables[0], PlantillaSoporteRojo.ToString()).ToString();
                }

                /*ORDENES DE TRABAJO*/

                if (listadoSopWosVerde.Tables[0].Rows.Count > 0)
                {

                    strBodySopWosVerde = Transform.TransformII.TransformXml(listadoSopWosVerde.Tables[0], PlantillaSoporteVerde.ToString()).ToString();
                }

                if (listadoSopWosNaranja.Tables[0].Rows.Count > 0)
                {
                    strBodySopWosNaranja = Transform.TransformII.TransformXml(listadoSopWosNaranja.Tables[0], PlantillaSoporteNaranja.ToString()).ToString();
                }

                if (listadoSopWosRojo.Tables[0].Rows.Count > 0)
                {
                    strBodySopWosRojo = Transform.TransformII.TransformXml(listadoSopWosRojo.Tables[0], PlantillaSoporteRojo.ToString()).ToString();
                }


                string ToMail = correosAnalista;
                if (ToMail.Replace(",", "").Trim().Length == 0)
                    ToMail = correosSupervisor;

                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                //service.Credentials = new WebCredentials(Dts.Variables["Matricula"].Value.ToString(), Dts.Variables["Password"].Value.ToString(), "BCPDOM");
               //service.Credentials = new WebCredentials(matricula, password, "BCPDOM");
                service.TraceEnabled = true;
                service.TraceFlags = TraceFlags.All;
                service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                service.Credentials = new WebCredentials(email, password);
                //service.AutodiscoverUrl(email);
                //service.AutodiscoverUrl("eboluarte@bcp.com.pe", RedirectionUrlValidationCallback);
                EmailMessage oMail = new EmailMessage(service);
                oMail.Subject = "[Notificación Automática] - Lista de INC/WO Activos";


                string bodyCompleto = "Buenos días Estimados, </br></br> Se les envía el reporte correspondiente a su Dominio </br></br> ";
                bodyCompleto += "Detalle de Incidentes:</br></br>" + strBodySopIncRojo + strBodySopIncNaranja + strBodySopIncVerde
                    + "</br></br>Detalle de Ordenes de Trabajos:</br></br>"
                    + strBodySopWosRojo + strBodySopWosNaranja + strBodySopWosVerde +
                    "</br></br>Quedamos pendiente ante cualquier consulta o comentario.</br>Gracias.</br>";


                // bodyCompleto += "</br></br><strong><span >RECORDAR:</span> En los nuevos indicadores las estimaciones DEBEN cerrarse en 6 días.<strong>";
                bodyCompleto += "</br>Saludos.</br>";

                oMail.Body = new MessageBody(bodyCompleto.ToString());
                oMail.Importance = Importance.High;


                string[] correos;
                correos = correosAnalista.Split(',');
                foreach (string correo in correos)
                    if (correo.Length > 0)
                        oMail.ToRecipients.Add(correo);


                string[] correosCc;
                correosCc = correosSupervisor.Split(',');
                foreach (string correoCc in correosCc)
                    if (correoCc.Length > 0)
                        oMail.CcRecipients.Add(correoCc);

                oMail.BccRecipients.Add(email);
                oMail.Body.BodyType = BodyType.HTML;



                _log.Info(new { Body = "Enviando Correo de Correo de Notificación Soporte Diario TCS" });

                oMail.Send();
                _log.Info(new { Body = "Finalizó Envio de  Correo de Notificación Soporte Diario TCS" });



            }
        }
        }
}
