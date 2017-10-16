using System;
using System.Collections.Generic;
using System.Web;
using System.Web.Services;
using System.Data;
using System.Data.SqlClient;
using System.Web.Script.Serialization;
using System.Web.Script.Services;
using Newtonsoft.Json;
using System.Net.Mail;
using System.Configuration;
using System.IO;
using System.Text;

namespace saf.views
{
    /// <summary>
    /// Descripción breve de capaservicio
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    [ScriptService]

    public class capaservicio : System.Web.Services.WebService
    {

        public class tabla_codificacion
        {
            public string nombre_tabla;
            public int codigo;
            public string descripcion;
            public string abrev;
            public int tipo_contenedor;

        }
        public class servicio{
            public String nombre_cliente;
            public long id_cliente;
        }
        public class formato
        {
            public String nombre_formato;
            public long id_formato;
            public string descripcionTipoDoc;
        }

        public class tipoEnvio
        {
            public string codigo;
            public string descripcion;
        }

        public class estadoSolicitud
        {
            public string codigo;
            public string descripcion;
        }

        public class tiempoRespuesta
        {
            public string codigo;
            public string descripcion;
        }
        public class tipoSalida
        {
            public string codigo;
            public string descripcion;
        }

        public class formaEnvio
        {
            public string codigo;
            public string descripcion;
        }
        public class MensajeRetorno
        {
            public string mensaje;
            public string fecha;
            public string hora;
        }

        public class usuario
        {
            public long id_usuario;
            public String nombre_usuario;
            public long id_unidad;
            public long id_perfil;
            public string email;
            public long estado_usuario;
            public long grupo_saf;
            public long id_cliente;
            public int usuarioTCS;
        }
        public class campoServicio
        {
            public String nombre_campo;
            public String salida;
            public int tipo_valor;
            public int largo;
            public int orden;
            public int tabla_codificacion;
            public string nombre_tabla;
        }


        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void clienteServicios(long idCliente, int usuarioTCS)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();

            SqlCommand cm = new SqlCommand("spGetServiciosUsuario @idCliente=" + idCliente.ToString() + ", @usuarioTCS=" +  usuarioTCS.ToString() , cn);

            SqlDataReader Dr = cm.ExecuteReader();
            List<servicio> lista = new List<servicio>();

            while (Dr.Read()){
                servicio s = new servicio();
                s.id_cliente = int.Parse( Dr["id_cliente"].ToString());
                s.nombre_cliente = Dr["nombre_cliente"].ToString();
                lista.Add(s);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void DatosUsuario(long idUsuario)
        {
            List<usuario> lista = new List<usuario>();
            usuario user = new usuario();
            string login = "";
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["sicprivado"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spGetUsuarios_ObtenerDatosUsuario @idUsuario=" + idUsuario.ToString(), cn);

            SqlDataReader Dr2 = cm.ExecuteReader();

            int tipoTCS = 0;
            if (Dr2.Read())
            {
                login = Dr2["login"].ToString();
           
                if (Dr2["perfil"].ToString().ToUpper() == "ADMINISTRADOR")
                {
                    tipoTCS = 1;
                }
            }
            cn.Close();

            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            cm.CommandText = "spGetUsuarioServicio @login='" + login +"'";
            SqlDataReader Dr = cm.ExecuteReader();

            if (Dr.Read())
            {
                user.id_usuario = long.Parse(Dr["id_usuario"].ToString());
                user.nombre_usuario = Dr["nombre_usuario"].ToString();
                user.id_unidad=long.Parse( Dr["id_unidad"].ToString());
                user.id_perfil = long.Parse(Dr["id_perfil"].ToString());
                user.email = Dr["email"].ToString();
                user.grupo_saf = long.Parse(Dr["grupo_saf"].ToString());
                user.id_cliente = long.Parse(Dr["id_cliente"].ToString());
                user.usuarioTCS = tipoTCS;
                user.estado_usuario= long.Parse(Dr["estado_usuario"].ToString());
                lista.Add(user);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void camposFormatoServicio(long idFormato)
        {
            List<campoServicio> lista = campos_Formato_Servicio(idFormato);
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write( jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void grabarSolicitudIngresoJSON(string idServicio, string idFormato, string Usuario, string idUsuario, string detalleJson)
        {
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));
            dt.Columns.RemoveAt(dt.Columns.Count - 1);
            lista =grabarSolicitudIngreso(idServicio, idFormato, Usuario, idUsuario, dt);
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public string grabarSolicitudIngresoDataSet(string idServicio, string idFormato, string Usuario, string idUsuario,  DataSet Ds)
        {
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = Ds.Tables[0];
            lista = grabarSolicitudIngreso(idServicio, idFormato, Usuario, idUsuario, dt);
            return jss.Serialize(lista);
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void crearSolicitudRecuperacionCajas_JSON(string idServicio, string idUsuario, string nombreUsuario, string observaciones, string tipoSolicitud, string tipoEnvio,
            string formaEnvio, string tipoSalida, string tiempoServicio, string detalleJson)
        {
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));
            lista = crearSolicitudRecuperacionCajas(idServicio, idUsuario, nombreUsuario, observaciones, tipoSolicitud, tipoEnvio, formaEnvio, tipoSalida, tiempoServicio, dt);
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void crearSolicitudRecuperacionDocumentos_JSON(string idServicio, string idFormato, string idUsuario, string nombreUsuario, string observaciones, string tipoSolicitud, string tipoEnvio,
        string formaEnvio, string tipoSalida, string tiempoServicio, string detalleJson)
        {
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));
            lista = crearSolicitudRecuperacionDocumentos(idServicio, idFormato, idUsuario, nombreUsuario, observaciones, tipoSolicitud, tipoEnvio, formaEnvio, tipoSalida, tiempoServicio, dt);
            Context.Response.Write(jss.Serialize(lista));
        }

        private List<MensajeRetorno> grabarSolicitudIngreso(string idServicio, string idFormato, string Usuario, string idUsuario, DataTable dt)
        {
            var registros = 0;
            registros = dt.Rows.Count;
            MensajeRetorno mensaje = new MensajeRetorno();

            List<campoServicio> campos = campos_Formato_Servicio(int.Parse(idFormato));
            List<MensajeRetorno> lista = new List<MensajeRetorno>();

            if (campos.Count != dt.Columns.Count)
            {
                mensaje.mensaje = "-1";
                lista.Add(mensaje);
                return lista;
            }

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spInsCabeceraSolicitudIngreso @idServicio=" + idServicio + ", @idFormato=" + idFormato + ", @idUsuario=" + idUsuario, cn);
            SqlDataReader Dr = cm.ExecuteReader();
            long idLote = 0;
            long idUnidad = 0;
            long idTipoDoc = 0;
            string nombreCliente = "";
            if (Dr.Read())
            {
                idLote = int.Parse(Dr["IDLote"].ToString());
                idUnidad = int.Parse(Dr["idUnidad"].ToString());
                idTipoDoc = int.Parse(Dr["idTipoDoc"].ToString());
                nombreCliente = Dr["nombreCliente"].ToString();
                mensaje.mensaje = idLote.ToString();
            }
            Dr.Close();

            string cadTipo = " id_TipoDoc, ";

            for (int c = 0; c <= campos.Count - 1; c++)
            {
                if (campos[c].nombre_campo.ToUpper() == "ID_TIPODOC")
                {
                    cadTipo = "";
                }
            }

            string cadena = "";
            string cadenaInventario = "";


            for (int c = 0; c <= campos.Count - 1; c++)
            {
                cadena += campos[c].nombre_campo + ", ";
                cadenaInventario += campos[c].nombre_campo + ", ";
            }




            cadena = cadena.Substring(0, cadena.Length - 2) + ") ";
            cadenaInventario = cadenaInventario.Substring(0, cadenaInventario.Length - 2) + ") ";
            string fecha = "";
            string cadPaso = "";
            string cadPasoOR = "";
            for (int r = 0; r <= dt.Rows.Count - 1; r++)
            {
                if (cadTipo == "")
                {
                    cadPaso = "";
                    cadPasoOR = "";
                }
                else
                {
                    cadPaso = "'" + idTipoDoc.ToString() + "', ";
                    cadPasoOR = idTipoDoc.ToString() + ", ";
                }

                string valores = "Values(" + idLote.ToString() + ", " + idServicio + ", " + idUnidad.ToString() + ", 0, 4, " + (r + 1).ToString() + ", " + cadPasoOR;

                string valoresInv = "Values(" + idLote.ToString() + ", " + idLote.ToString() + ", " + idLote.ToString() + ", " + idServicio + ", " + idUnidad.ToString() + ", 0, 4, " +
                    idTipoDoc.ToString() + ", '" + Usuario.ToString() + "', convert(varchar(10),getdate(),112), " + "1, " + cadPaso;

                for (int c = 0; c <= dt.Columns.Count - 1; c++)
                {
                    if (campos[c].tipo_valor == 1)
                    {
                        valores += "'" + dt.Rows[r][c].ToString() + "', ";
                        valoresInv += "'" + dt.Rows[r][c].ToString() + "', ";
                    }
                    if (campos[c].tipo_valor == 2)
                    {
                        valores += dt.Rows[r][c].ToString() + ", ";
                        valoresInv += "'" + dt.Rows[r][c].ToString() + "', ";
                    }
                    if (campos[c].tipo_valor == 3)
                    {
                        if (campos[c].tabla_codificacion == 1)
                        {
                            valores += "'" + dt.Rows[r][c].ToString() + "', ";
                            valoresInv += "'" + dt.Rows[r][c].ToString() + "', ";
                        }
                        else
                        {
                            fecha = dt.Rows[r][c].ToString().Replace("-", "").Replace("/", "");
                            fecha = fecha.Substring(4, 4) + fecha.Substring(2, 2) + fecha.Substring(0, 2);
                            valores += fecha.ToString() + ", ";
                            valoresInv += fecha.ToString() + ", ";
                        }

                    }
                }

                /* cm.CommandText = cadena + valores.Substring(0, valores.Length - 2) + ");" +
                                  cadenaInventario + valoresInv.Substring(0, valoresInv.Length - 2) + ");";
 */
                valores = valores.Substring(0, valores.Length - 2);
                valoresInv = valoresInv.Substring(0, valoresInv.Length - 2);
                cm.CommandText = "spInsDetIngresoCaja @cadTipo='" + cadTipo + "', @cadenaValores='" +cadena + valores.Replace("'", "''") + "', @cadenaInventario='" + cadenaInventario+ valoresInv.Replace("'", "''") + "'";
                cm.ExecuteNonQuery();
            }

            cm.CommandText = "spActualizar_IngresoCaja_Saf_Caja @idLote="+idLote.ToString();
            cm.ExecuteNonQuery();
            cm.CommandText = "spDelOrdenTemporal @idCliente=" + idServicio + ", @idFormato=" + idFormato + ", @idUsuario=" + idUsuario + ", @tipoSolicitud=1";
            cm.ExecuteNonQuery();
            cn.Close();
            mensaje.mensaje = idLote.ToString();
            lista.Add(mensaje);

            EnviarCorreoSolicitudIngreso(idLote, nombreCliente, Usuario);

            return lista;
        }

        private void EnviarCorreoSolicitudIngreso(long idLote, string nombreCliente, string nombreUsuario)
        {
            string body = "Una nueva solicitud del cliente-servicio <b>" + nombreCliente + "</b> se ha generado.</br></br>" +
            "<table border='1' cellspacing='0'>" +
                "<tr><td width='300px'><h3 class='text-info'>Número de Solicitud</h3></td><td width='150px'><h3 class='text-info'><b>" + idLote.ToString() + "</b></h3></td></tr>" +
                "<tr><td width='300px'><h4 class='text-info'>Fecha creación </h4></td><td width='150px'><h4 class='text-info'>" + DateTime.Today.ToShortDateString() + "</h4></td></tr>" +
                "<tr><td width='300px'><h4 class='text-info'>Hora creación </h4></td><td width='150px'><h4 class='text-info'>" + DateTime.Now.ToString("HH:mm") + "</h4></td></tr>" +
            "</table>";
            
            MailMessage email = new MailMessage();

            string correorDestino=ConfigurationManager.AppSettings.Get("correoDestinatarioSolicitud");

            char[] delimitador = new char[] { ';' };
            if (correorDestino != "")
            {
                foreach (string copiar_a in correorDestino.Split(delimitador))
                {
                    email.To.Add(new MailAddress(copiar_a));
                }
            }

            email.From = new MailAddress(ConfigurationManager.AppSettings.Get("correoEmisor"));
            email.Subject = "Solicitud de Ingreso de Cajas";
            email.Body = body;
            email.IsBodyHtml = true;
            email.Priority = MailPriority.Normal;

            SmtpClient smtp = new SmtpClient();
            smtp.Host =ConfigurationManager.AppSettings.Get("hostSmtp");
            smtp.Port = int.Parse( ConfigurationManager.AppSettings.Get("puertoSmtp"));
            if (ConfigurationManager.AppSettings.Get("habilitarSsl")=="true")
            {
                smtp.EnableSsl = true;
            }
            else
            {
                smtp.EnableSsl = false;
            }
            if (ConfigurationManager.AppSettings.Get("usarCredencialesporDefault") == "true")
            {
                smtp.UseDefaultCredentials = true;
            }
            else
            {
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings.Get("correoUsuario"), ConfigurationManager.AppSettings.Get("correoContrasena"));
            }

            string output = null;

            try
            {
                smtp.Send(email);
                email.Dispose();
                output = "Corre electrónico fue enviado satisfactoriamente.";
            }
            catch (Exception ex)
            {
                output = "Error enviando correo electrónico: " + ex.Message;
            }
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void servicioFormatos(long idServicio, int tipoSolicitud)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spGetClienteFormatosEntrada @idcliente=" + idServicio.ToString()+", @tipoProceso="+tipoSolicitud.ToString(), cn);
            SqlDataReader Dr = cm.ExecuteReader();
            List<formato> lista = new List<formato>();

            while (Dr.Read())
            {
                formato s = new formato();
                s.id_formato = int.Parse(Dr["id_formato"].ToString());
                s.nombre_formato = Dr["nombre_formato"].ToString();
                s.descripcionTipoDoc = Dr["descripcionTipoDoc"].ToString();
                lista.Add(s);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        public bool SaveDocument(Byte[] docbinaryarray)
        {
            string strdocPath;
            strdocPath = "C:\\TempFiles\\prueba.doc";
            FileStream objfilestream = new FileStream(strdocPath, FileMode.Create, FileAccess.ReadWrite);
            objfilestream.Write(docbinaryarray, 0, docbinaryarray.Length);
            objfilestream.Close();
            return true;
        }

        [WebMethod]
        public void Guardar()
        {
            HttpContext Contexto = HttpContext.Current;
            HttpFileCollection ColeccionArchivos = Context.Request.Files;
            MensajeRetorno mensaje = new MensajeRetorno();
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            String NombreArchivo = "";
            for (int ArchivoActual = 0; ArchivoActual < ColeccionArchivos.Count; ArchivoActual++)
            {
                NombreArchivo = ColeccionArchivos[ArchivoActual].FileName;
                ///  string DatosArchivo = System.IO.Path.GetFileName(ColeccionArchivos[ArchivoActual].FileName);
                ///  string CarpetaParaGuardar = Server.MapPath("Archivos") + "\\" + DatosArchivo;
                ///  ColeccionArchivos[ArchivoActual].SaveAs(CarpetaParaGuardar);
                 mensaje.mensaje = NombreArchivo;
            }
            lista.Add(mensaje);
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
         }


        private List<campoServicio> campos_Formato_Servicio(long idFormato) 
        {
            List<campoServicio> lista = new List<campoServicio>();
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spGetCamposFormatoEntrada @idFormato=" + idFormato.ToString(), cn);

            SqlDataReader Dr = cm.ExecuteReader();

            while (Dr.Read())
            {
                campoServicio campo = new campoServicio();
                campo.nombre_campo = Dr["CampoNombreSAFInventario"].ToString();
                campo.salida = Dr["CampoNombreCliente"].ToString();
                campo.tipo_valor = int.Parse(Dr["tipo_valor"].ToString());
                campo.largo = int.Parse(Dr["largo"].ToString());
                campo.tabla_codificacion = int.Parse(Dr["TipoTablaInv"].ToString());
                campo.nombre_tabla = Dr["nombre_tabla"].ToString();
                lista.Add(campo);
            }
            cn.Close();
            return(lista);
        }

        [WebMethod]
        public DataSet campos_Formato_Servicio_DataTable(long idFormato)
        {
            DataSet Ds = new DataSet();
            using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString))
            {
                string ConsultaProductos = "spGetCamposFormatoEntrada @idFormato=" + idFormato.ToString();
                SqlCommand cmd = new SqlCommand(ConsultaProductos, cnn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(Ds);
            }
           
            return Ds;
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void guardarSolicitudIngresoTMP(string idServicio, string idFormato, string idUsuario, string detalleJson, string tipoSolicitud)
        {
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));
            if (tipoSolicitud.ToString() == "1")
            {
                dt.Columns.RemoveAt(dt.Columns.Count - 1);
            }

            lista = guardarSolicitudIngreso(idServicio, idFormato, idUsuario, dt, tipoSolicitud);
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void guardarSolicitudRecuperacionCajas_TMP(string idServicio, string idUsuario, string detalleJson)
        {
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));

            lista = guardarSolicitudRecuperacionCajas(idServicio, idUsuario, dt);
            Context.Response.Write(jss.Serialize(lista));
        }


        private List<MensajeRetorno> guardarSolicitudIngreso(string idServicio, string idFormato, string idUsuario, DataTable dt, string tipoSolicitud)
        {
            var registros = 0;
            registros = dt.Rows.Count;
            MensajeRetorno mensaje = new MensajeRetorno();

            List<campoServicio> campos = campos_Formato_Servicio(int.Parse(idFormato));
            List<MensajeRetorno> lista = new List<MensajeRetorno>();



            if (tipoSolicitud == "1")
            {
                if (campos.Count != dt.Columns.Count)
                {
                    mensaje.mensaje = "-1";
                    lista.Add(mensaje);
                    return lista;
                }

            }
            else
            {
                campoServicio campoBodega = new campoServicio();
                campoBodega.nombre_campo = "Bodega";
                campoBodega.salida = "Bodega";
                campoBodega.tipo_valor = 1;
                campoBodega.largo = 100;
                campos.Add(campoBodega);

                campoServicio campoBandeja = new campoServicio();
                campoBandeja.nombre_campo = "Bandeja";
                campoBandeja.salida = "Bandeja";
                campoBandeja.tipo_valor = 1;
                campoBandeja.largo = 100;
                campos.Add(campoBandeja);

                campoServicio id_registro = new campoServicio();
                id_registro.nombre_campo = "id_registro";
                id_registro.salida = "id_registro";
                id_registro.tipo_valor = 2;
                id_registro.largo = 100;
                campos.Add(id_registro);
            }


            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spInsCabeceraSolicitudIngresoTemporal @idServicio=" + idServicio + ", @idFormato=" + idFormato + ", @idusuario=" +
                idUsuario + ", @tipoSolicitud=" + tipoSolicitud.ToString(), cn);
            SqlDataReader Dr = cm.ExecuteReader();
            long idLote = 0;
            long idUnidad = 0;
            long idTipoDoc = 0;
            if (Dr.Read())
            {
                idLote = int.Parse(Dr["IDLote"].ToString());
                idUnidad = int.Parse(Dr["idUnidad"].ToString());
                idTipoDoc = int.Parse(Dr["idTipoDoc"].ToString());
                mensaje.mensaje = idLote.ToString();
            }
            Dr.Close();

            string comando = "";
            string cadena = "";
            int agregaTipoDoc = 1;
            for (int c = 0; c <= campos.Count - 1; c++)
            {
                cadena += campos[c].nombre_campo + ", ";
                if (campos[c].nombre_campo.ToUpper() == "ID_TIPODOC")
                {
                    agregaTipoDoc = 0;
                }
            }

            cadena = cadena.Substring(0, cadena.Length - 2) + ") ";


            for (int r = 0; r <= dt.Rows.Count - 1; r++)
            {
                string valores = "Values(" + idLote.ToString() + ", " + idServicio + ", " + idUnidad.ToString() + ", 0, 1, " + (r + 1).ToString() + ", ";
                if (agregaTipoDoc==1)
                {
                    valores += idTipoDoc.ToString() + ", ";
                }
                for (int c = 0; c <= dt.Columns.Count - 1; c++)
                {
                    if (campos[c].tipo_valor == 1)
                    {
                        valores += "'" + dt.Rows[r][c].ToString() + "', ";
                    }
                    if (campos[c].tipo_valor == 2)
                    {
                        valores += dt.Rows[r][c].ToString() + ", ";
                    }
                    if (campos[c].tipo_valor == 3)
                    {
                        valores += "'" + dt.Rows[r][c].ToString() + "', ";
                    }
                }
                comando = cadena + valores.Substring(0, valores.Length - 2) + ");";
                cm.CommandText = "spInsDetIngresoTemporal @agregaTipoDoc=" + agregaTipoDoc.ToString() + ", @cadenaValores='"+comando.Replace("'", "''")+"'";
                cm.ExecuteNonQuery();
            }

            cm.CommandText = "spActualizar_IngresoCaja_Saf_Caja @idLote=" + idLote.ToString();
            cm.ExecuteNonQuery();

            cn.Close();
            mensaje.mensaje = idLote.ToString();
            lista.Add(mensaje);
            return lista;
        }

        private List<MensajeRetorno> guardarSolicitudRecuperacionCajas(string idServicio, string idUsuario, DataTable dt)
        {
            var registros = 0;
            registros = dt.Rows.Count;
            MensajeRetorno mensaje = new MensajeRetorno();
            List<MensajeRetorno> lista = new List<MensajeRetorno>();

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spInsCabeceraSolicitudIngresoTemporal @idServicio=" + idServicio + ", @idFormato=0, @idusuario=" +
                idUsuario + ", @tipoSolicitud=3", cn);
            SqlDataReader Dr = cm.ExecuteReader();
            long idLote = 0;
            long idUnidad = 0;
            long idTipoDoc = 0;
            if (Dr.Read())
            {
                idLote = int.Parse(Dr["IDLote"].ToString());
                idUnidad = int.Parse(Dr["idUnidad"].ToString());
                idTipoDoc = int.Parse(Dr["idTipoDoc"].ToString());
                mensaje.mensaje = idLote.ToString();
            }
            Dr.Close();

            string comando = "";
            for (int r = 0; r <= dt.Rows.Count - 1; r++)
            {
                comando = "spInsTempRecuperaCaja @idLote="+ idLote.ToString() +", @idCaja=" + dt.Rows[r][0].ToString();
                cm.CommandText = comando;
                cm.ExecuteNonQuery();
            }

            cn.Close();
            mensaje.mensaje = idLote.ToString();
            lista.Add(mensaje);
            return lista;
        }





        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void DatosTemporales(long idFormato, long idUsuario, string tipoSolicitud)
        {
            JavaScriptSerializer jss = new JavaScriptSerializer();

            
            DataTable Dt = new DataTable();
            DataSet Ds = new DataSet();
            DataTable dt2 = new DataTable();
            DataTable dt = new DataTable();
            string otrosCampos = "";
            if (tipoSolicitud.ToString()!="1")
            {
               otrosCampos=", SAF_OR_Detalle_TMP.Bodega, SAF_OR_Detalle_TMP.Bandeja, SAF_OR_Detalle_TMP.id_registro";
            }
            using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString))
            {
                string ConsultaProductos = "spGetCamposFormatoEntrada @idFormato=" + idFormato.ToString();
                SqlCommand cmd = new SqlCommand(ConsultaProductos, cnn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt2);

                string cadena = "select SAF_OR_Detalle_TMP.idTemporal, ";

                for (int x = 0; x <= dt2.Rows.Count - 1; x++)
                {
                    cadena = cadena + "SAF_OR_Detalle_TMP." + dt2.Rows[x]["campoNombreSafInventario"].ToString() + ", ";
                }

                cadena = cadena.Substring(0, cadena.Length - 2);
                cadena = cadena + otrosCampos + " from SAF_OR_Detalle_TMP inner join SAF_OR_LOTE_TMP on ( SAF_OR_LOTE_TMP.Id_Lote_OR=SAF_OR_Detalle_TMP.Id_Lote_OR) where " +
                    " SAF_OR_LOTE_TMP.Id_Formato=" +idFormato.ToString() + " and SAF_OR_LOTE_TMP.id_Usuario=" + idUsuario.ToString() +
                    " and SAF_OR_LOTE_TMP.tipoSolicitud=" + tipoSolicitud.ToString()+ " order by id_registro";

                cmd.CommandText = cadena;
                SqlDataAdapter d2 = new SqlDataAdapter(cmd);
                d2.Fill(dt);


            }
            Context.Response.Write(ConvertDataTableTojSonString(dt));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void getRecuperacionCajasTemporal(long idUsuario, long idCliente)
        {
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = new DataTable();

            using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString))
            {
                string cadena = "spGetRecuperacionCajasTemporal @idCliente="+idCliente.ToString()+", @idUsuario="+idUsuario.ToString();
                SqlCommand cmd = new SqlCommand(cadena, cnn);
                SqlDataAdapter d2 = new SqlDataAdapter(cmd);
                d2.Fill(dt);
            }
            Context.Response.Write(ConvertDataTableTojSonString(dt));
        }


        public String ConvertDataTableTojSonString(DataTable dataTable)
        {
            System.Web.Script.Serialization.JavaScriptSerializer serializer =
                   new System.Web.Script.Serialization.JavaScriptSerializer();

            List<Dictionary<String, Object>> tableRows = new List<Dictionary<String, Object>>();

            Dictionary<String, Object> row;

            foreach (DataRow dr in dataTable.Rows)
            {
                row = new Dictionary<String, Object>();
                foreach (DataColumn col in dataTable.Columns)
                {
                    row.Add(col.ColumnName, dr[col]);
                }
                tableRows.Add(row);
            }
            return serializer.Serialize(tableRows);
        }


        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void tipoEnvioClienteServicio(long idServicio)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spGetTipoEnvioClienteServicio @idcliente=" + idServicio.ToString(), cn);
            SqlDataReader Dr = cm.ExecuteReader();
            List<tipoEnvio> lista = new List<tipoEnvio>();

            while (Dr.Read())
            {
                tipoEnvio s = new tipoEnvio();
                s.codigo = Dr["codigo"].ToString();
                s.descripcion = Dr["descripcion"].ToString();
                lista.Add(s);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void tiempoRespuestaClienteServicio(long idServicio)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spGetTiempoRespuestaClienteServicio @idcliente=" + idServicio.ToString(), cn);
            SqlDataReader Dr = cm.ExecuteReader();
            List<tiempoRespuesta> lista = new List<tiempoRespuesta>();

            while (Dr.Read())
            {
                tiempoRespuesta s = new tiempoRespuesta();
                s.codigo = Dr["codigo"].ToString();
                s.descripcion = Dr["descripcion"].ToString();
                lista.Add(s);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void tipoDocumentoCliente(long idServicio)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("aqpSAF_TipoDocumento @idcliente=" + idServicio.ToString(), cn);
            SqlDataReader Dr = cm.ExecuteReader();
            List<tiempoRespuesta> lista = new List<tiempoRespuesta>();

            while (Dr.Read())
            {
                tiempoRespuesta s = new tiempoRespuesta();
                s.codigo = Dr["id_tipodoc"].ToString();
                s.descripcion = Dr["descripcion"].ToString();
                lista.Add(s);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }


        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void formadeEnvio()
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spGetFormaEnvio", cn);
            SqlDataReader Dr = cm.ExecuteReader();
            List<formaEnvio> lista = new List<formaEnvio>();

            while (Dr.Read())
            {
                formaEnvio s = new formaEnvio();
                s.codigo = Dr["codigo"].ToString();
                s.descripcion = Dr["descripcion"].ToString();
                lista.Add(s);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void tipodeSalida()
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spGetTipoSalida", cn);
            SqlDataReader Dr = cm.ExecuteReader();
            List<tipoSalida> lista = new List<tipoSalida>();

            while (Dr.Read())
            {
                tipoSalida s = new tipoSalida();
                s.codigo = Dr["codigo"].ToString();
                s.descripcion = Dr["descripcion"].ToString();
                lista.Add(s);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void buscarCajas(string idServicio, string detalleJson)
        {
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));


            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            DataTable dt2 = new DataTable();
            string ConsultaProductos = "";
            foreach (DataRow dr in dt.Rows)
            {
                ConsultaProductos = "spGetCajas @idCliente=" + idServicio.ToString() + ", @desde=" + dr[2].ToString().Trim() + ", @hasta=" + dr[3].ToString().Trim();
            }
            SqlCommand cmd = new SqlCommand(ConsultaProductos, cn);
            SqlDataAdapter d2 = new SqlDataAdapter(cmd);
            d2.Fill(dt2);
            cn.Close();
            Context.Response.Write(ConvertDataTableTojSonString(dt2));
        }


        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void consultarRegistroPorTipoDocumento(string idServicio, string idFormato, string detalleJson)
        {
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));
           // dt.Columns.RemoveAt(dt.Columns.Count - 1);
            string cadena = "id_cliente=" + idServicio +" and " ;
            foreach (DataRow dr in dt.Rows)
            {
                if (dr[1].ToString() == "numerico")
                {
                    cadena = cadena + "(" + dr[0].ToString() + " between '" + dr[2].ToString().Trim() + "' and '" + dr[3].ToString().Trim() + "') and ";
                }
                if (dr[1].ToString() == "texto")
                {
                    cadena = cadena + "(" + dr[0].ToString() + "='" + dr[2].ToString() + "') and ";
                }
                if (dr[1].ToString() == "fecha")
                {
                    cadena = cadena + "(" + dr[0].ToString() + " between '" + dr[2].ToString() + "' and '" + dr[3].ToString() + "') and ";
                }
            }

            if (cadena.Length > 0)
            {
                cadena = cadena.Substring(0, cadena.Length - 4);
            }

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            DataTable dt1 = new DataTable();
            string ConsultaProductos = "spGetCamposFormatoEntrada @idFormato=" + idFormato.ToString();
            SqlCommand cmd = new SqlCommand(ConsultaProductos, cn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt1);

            string campos = "select SAF_Inventario.id_registro, SAF_Inventario.ID_estado, ";

            for (int x = 0; x <= dt1.Rows.Count - 1; x++)
            {
                campos = campos + "SAF_Inventario." + dt1.Rows[x]["campoNombreSafInventario"].ToString() + ", ";
            }

            campos = campos + "Coalesce(SAF_Inventario.bodega,'') as Bodega, Coalesce(SAF_Inventario.bandeja,'') as Bandeja, SAF_Codificacion.Descripcion as Estado,  SUBSTRING(ltrim(Saf_inventario.Fecha_estado),7,2)+'/'+SUBSTRING(ltrim(Saf_inventario.Fecha_estado),5,2)+'/'+SUBSTRING(ltrim(Saf_inventario.Fecha_estado),1,4) as fechaEstado , Coalesce(id_lote_solicitud,0) as idLoteSolicitud, coalesce(solicitante,'') as solicitante";

            DataTable dt2 = new DataTable();
            cmd.Connection = cn;
            campos = campos.Substring(0, campos.Length - 2);

            cadena = campos + " from SAF_Inventario inner join SAF_Codificacion on ( SAF_Codificacion.nombre_tabla='ESTADO_INV' and " + 
                "SAF_Codificacion.Codigo=SAF_Inventario.Id_Estado) where " + cadena;

            cmd.CommandText = cadena;
            SqlDataAdapter d2 = new SqlDataAdapter(cmd);
            d2.Fill(dt2);
            cn.Close();
            Context.Response.Write(ConvertDataTableTojSonString(dt2));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void grabarSolicitudRecuperacionJSON(string idServicio, string idFormato, string idUsuario, string nombreUsuario, string observaciones, string tipoSolicitud, string tipoEnvio, 
            string formaEnvio, string tipoSalida, string tiempoServicio, string detalleJson)
        {
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));
            lista = grabarSolicitudRecuperacion(idServicio, idFormato, idUsuario, nombreUsuario, observaciones, tipoSolicitud, tipoEnvio, formaEnvio, tipoSalida, tiempoServicio, dt);
            Context.Response.Write(jss.Serialize(lista));
        }

        private List<MensajeRetorno> grabarSolicitudRecuperacion(string idServicio, string idFormato, string idUsuario,string nombreUsuario, string observaciones, string tipoSolicitud, string tipoEnvio,
            string formaEnvio, string tipoSalida, string tiempoServicio, DataTable dt)
        {
            var registros = 0;
            registros = dt.Rows.Count;
            MensajeRetorno mensaje = new MensajeRetorno();
            List<MensajeRetorno> lista = new List<MensajeRetorno>();

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spInsCabeceraSolicitudRecuperacion @idServicio=" + idServicio + ", @idFormato=" + idFormato + ", @idUsuario=" + idUsuario + 
                ", @tipoSolicitud="+tipoSolicitud+", @tipoEnvio=" + tipoEnvio + ", @formaEnvio="+formaEnvio+", @tipoSalida="+tipoSalida+", @tiempoServicio=" + tiempoServicio + 
                ", @observaciones='"+observaciones+"'", cn);

            SqlDataReader Dr = cm.ExecuteReader();
            long idSolicitud = 0;
            string nombreCliente = "";
            if (Dr.Read())
            {
                idSolicitud = int.Parse(Dr["idSolicitud"].ToString());
                nombreCliente = Dr["nombreCliente"].ToString();
                mensaje.mensaje = idSolicitud.ToString();
                mensaje.fecha = Dr["fecha"].ToString();
                mensaje.hora = Dr["hora"].ToString();
            }
            Dr.Close();


            for (int r = 0; r <= dt.Rows.Count - 1; r++)
            {
                string valores = idSolicitud.ToString() + ", " + dt.Rows[r][0].ToString() ;

                cm.CommandText = "spInsDetalleSolicitudRecuperacionDocumentos @cadena='" + valores + "'";
                cm.ExecuteNonQuery();
            }
            lista.Add(mensaje);
            int tipoSol = 0;

            tipoSol = int.Parse(tipoSolicitud) + 1;

            cm.CommandText = "spDelOrdenTemporal @idCliente=" + idServicio + ", @idFormato=" + idFormato + ", @idUsuario=" + idUsuario + ", @tipoSolicitud="+tipoSol.ToString();
            cm.ExecuteNonQuery();
            cn.Close();

            EnviarCorreoSolicitudRecuperacionDocumentos(idSolicitud, nombreCliente, nombreUsuario, tipoSolicitud);

            return lista;
        }


        private List<MensajeRetorno> crearSolicitudRecuperacionCajas(string idServicio, string idUsuario, string nombreUsuario, string observaciones, string tipoSolicitud, string tipoEnvio,
            string formaEnvio, string tipoSalida, string tiempoServicio, DataTable dt)
        {
            var registros = 0;
            registros = dt.Rows.Count;
            MensajeRetorno mensaje = new MensajeRetorno();
            List<MensajeRetorno> lista = new List<MensajeRetorno>();

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spInsCabaceraSolicitudRecuperacionCajas @idServicio=" + idServicio + ", @idUsuario=" + idUsuario +
                ", @tipoSolicitud=" + tipoSolicitud + ", @tipoEnvio=" + tipoEnvio + ", @formaEnvio=" + formaEnvio + ", @tipoSalida=" + tipoSalida + ", @tiempoServicio=" + tiempoServicio +
                ", @observaciones='" + observaciones + "',@nombreUsuario='"+nombreUsuario.ToString()+"'", cn);

            SqlDataReader Dr = cm.ExecuteReader();
            long idSolicitud = 0;
            string idUnidad = "";
            string nombreCliente = "";
            string idTipoDoc = "";
            if (Dr.Read())
            {
                idSolicitud = int.Parse(Dr["correlativo"].ToString());
                nombreCliente = Dr["nombreCliente"].ToString();
                idUnidad= Dr["idUnidad"].ToString();
                idTipoDoc = Dr["idTipoDoc"].ToString();
                mensaje.mensaje = idSolicitud.ToString();
                mensaje.fecha = Dr["fecha"].ToString();
                mensaje.hora = Dr["hora"].ToString();
            }
            Dr.Close();


            for (int r = 0; r <= dt.Rows.Count - 1; r++)
            {
                string valores = idSolicitud.ToString() + ", " + dt.Rows[r][0].ToString();

                cm.CommandText = "spInsDetalleSolicitudRecuperacionCajas @idLote=" + idSolicitud.ToString() + ", @idCaja=" + dt.Rows[r][0].ToString() + 
                    ", @idCliente=" + idServicio.ToString() + ", @idUnidad=0, @nombreUsuario='" + nombreUsuario.ToString() + "', @idTipoDoc="+idTipoDoc.ToString();

                cm.ExecuteNonQuery();
            }
            lista.Add(mensaje);
            int tipoSol = 0;

            tipoSol = int.Parse(tipoSolicitud) + 1;

            cm.CommandText = "spDelOrdenTemporal @idCliente=" + idServicio + ", @idUsuario=" + idUsuario + ", @tipoSolicitud=" + tipoSol.ToString();
            cm.ExecuteNonQuery();
            cn.Close();

            EnviarCorreoSolicitudRecuperacionDocumentos(idSolicitud, nombreCliente, nombreUsuario, tipoSolicitud);

            return lista;
        }

        private List<MensajeRetorno> crearSolicitudRecuperacionDocumentos(string idServicio, string idFormato, string idUsuario, string nombreUsuario, string observaciones, string tipoSolicitud, string tipoEnvio,
            string formaEnvio, string tipoSalida, string tiempoServicio, DataTable dt)
        {
            var registros = 0;
            registros = dt.Rows.Count;
            MensajeRetorno mensaje = new MensajeRetorno();
            List<MensajeRetorno> lista = new List<MensajeRetorno>();

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand("spInsCabaceraSolicitudRecuperacionDocumentos @idServicio=" + idServicio + ", @idFormato="+ idFormato.ToString() +", @idUsuario=" + idUsuario +
                ", @tipoSolicitud=" + tipoSolicitud + ", @tipoEnvio=" + tipoEnvio + ", @formaEnvio=" + formaEnvio + ", @tipoSalida=" + tipoSalida + ", @tiempoServicio=" + tiempoServicio +
                ", @observaciones='" + observaciones + "', @nombreUsuario='"+nombreUsuario.ToString()+"'", cn);

            SqlDataReader Dr = cm.ExecuteReader();
            long idSolicitud = 0;
            string idUnidad = "";
            string nombreCliente = "";
            string idTipoDoc = "";
            if (Dr.Read())
            {
                idSolicitud = int.Parse(Dr["correlativo"].ToString());
                nombreCliente = Dr["nombreCliente"].ToString();
                idUnidad = Dr["idUnidad"].ToString();
                idTipoDoc = Dr["idTipoDoc"].ToString();
                mensaje.mensaje = idSolicitud.ToString();
                mensaje.fecha = Dr["fecha"].ToString();
                mensaje.hora = Dr["hora"].ToString();
            }
            Dr.Close();
            // Aqui se debe agregar el codigo para grabar el detalle de la solicitud de recuperacion de documentos

            for (int r = 0; r <= dt.Rows.Count - 1; r++)
            {
                cm.CommandText = "spInsDetalleRecuperacionDocumentos @idLote=" + idSolicitud.ToString() + ", @idRegistro=" + dt.Rows[r][0].ToString() +
                    ", @nombreUsuario='" + nombreUsuario.ToString() + "'";

                cm.ExecuteNonQuery();
            }
            lista.Add(mensaje);
            int tipoSol = 0;

            tipoSol = int.Parse(tipoSolicitud) + 1;

            cn.Close();

            EnviarCorreoSolicitudRecuperacionDocumentos(idSolicitud, nombreCliente, nombreUsuario, tipoSolicitud);

            return lista;
        }


        private void EnviarCorreoSolicitudRecuperacionDocumentos(long idLote, string nombreCliente, string nombreUsuario, string tipoSolicitud)
        {
            string tipo = "";
            if (tipoSolicitud == "1")
            {
                tipo = "Solicitud de Recuperación de Documentos";
            }
            else
            {
                tipo = "Solicitud de Recuperación de Cajas";
            }
            string body = "Una nueva " + tipo + " del cliente-servicio <b>" + nombreCliente + "</b> se ha generado.</br></br>" +
            "<table border='1' cellspacing='0'>" +
                "<tr><td width='300px'><h3 class='text-info'>Número de Solicitud</h3></td><td width='150px'><h3 class='text-info'><b>" + idLote.ToString() + "</b></h3></td></tr>" +
                "<tr><td width='300px'><h4 class='text-info'>Fecha creación </h4></td><td width='150px'><h4 class='text-info'>" + DateTime.Today.ToShortDateString() + "</h4></td></tr>" +
                "<tr><td width='300px'><h4 class='text-info'>Hora creación </h4></td><td width='150px'><h4 class='text-info'>" + DateTime.Now.ToString("HH:mm") + "</h4></td></tr>" +
            "</table>";

            MailMessage email = new MailMessage();
            string correorDestino = ConfigurationManager.AppSettings.Get("correoDestinatarioSolicitud");

            char[] delimitador = new char[] { ';' };
            if (correorDestino != "")
            {
                foreach (string copiar_a in correorDestino.Split(delimitador))
                {
                    email.To.Add(new MailAddress(copiar_a));
                }
            }

            email.From = new MailAddress(ConfigurationManager.AppSettings.Get("correoEmisor"));
            email.Subject = tipo;
            email.Body = body;
            email.IsBodyHtml = true;
            email.Priority = MailPriority.Normal;

            SmtpClient smtp = new SmtpClient();
            smtp.Host = ConfigurationManager.AppSettings.Get("hostSmtp");
            smtp.Port = int.Parse(ConfigurationManager.AppSettings.Get("puertoSmtp"));
            if (ConfigurationManager.AppSettings.Get("habilitarSsl") == "true")
            {
                smtp.EnableSsl = true;
            }
            else
            {
                smtp.EnableSsl = false;
            }
            if (ConfigurationManager.AppSettings.Get("usarCredencialesporDefault") == "true")
            {
                smtp.UseDefaultCredentials = true;
            }
            else
            {
                smtp.UseDefaultCredentials = false;
                smtp.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings.Get("correoUsuario"), ConfigurationManager.AppSettings.Get("correoContrasena"));
            }

            string output = null;

            try
            {
                smtp.Send(email);
                email.Dispose();
                output = "Corre electrónico fue enviado satisfactoriamente.";
            }
            catch (Exception ex)
            {
                output = "Error enviando correo electrónico: " + ex.Message;
            }
        }

        [WebMethod]
        public DataSet ServicioCamposClave(long idServicio)
        {
            DataSet Ds = new DataSet();
            using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString))
            {
                string Consulta = "spGetCampoClaveCliente @idCliente=" + idServicio.ToString();
                SqlCommand cmd = new SqlCommand(Consulta, cnn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(Ds);
            }
            return Ds;
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void SolicitudesporClienteServicio(string idServicio, string idEstado, string tipoSolicitud, string fechaDesde, string fechaHasta, string numeroSolicitud)
        {
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = new DataTable();

            if (numeroSolicitud == "")
            {
                numeroSolicitud = "0";
            }

            if( fechaDesde != "")
            {
                fechaDesde = fechaDesde.Substring(6, 4) + fechaDesde.Substring(3, 2) + fechaDesde.Substring(0, 2);
            }
            if (fechaHasta != "")
            {
                fechaHasta = fechaHasta.Substring(6, 4) + fechaHasta.Substring(3, 2) + fechaHasta.Substring(0, 2);
            }
            using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString))
            {
                SqlCommand cmd = new SqlCommand("spGetSolicitudPorClienteServicio @idCliente=" + idServicio+", @idEstado=" +idEstado + ", @tipoSolicitud=" + tipoSolicitud + 
                    ", @fechaDesde='" + fechaDesde + "', @fechaHasta='"+fechaHasta+"', @numeroSolicitud="+numeroSolicitud , cnn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
            }
            Context.Response.Write(ConvertDataTableTojSonString(dt));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void listaEstadosSolicitud()
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();

            SqlCommand cm = new SqlCommand("spGetEstadosSolicitud", cn);

            SqlDataReader Dr = cm.ExecuteReader();
            List<estadoSolicitud> lista = new List<estadoSolicitud>();

            while (Dr.Read())
            {
                estadoSolicitud s = new estadoSolicitud();
                s.codigo = Dr["codigo"].ToString();
                s.descripcion = Dr["descripcion"].ToString();
                lista.Add(s);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void codificacion()
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();

            SqlCommand cm = new SqlCommand("spGetCodificacion", cn);

            SqlDataReader Dr = cm.ExecuteReader();
            List<tabla_codificacion> lista = new List<tabla_codificacion>();

            while (Dr.Read())
            {
                tabla_codificacion item = new tabla_codificacion();
                item.nombre_tabla = Dr["nombre_tabla"].ToString();
                item.codigo = int.Parse( Dr["codigo"].ToString());
                item.descripcion = Dr["descripcion"].ToString();
                item.abrev = Dr["abrev"].ToString();
                item.tipo_contenedor = int.Parse(Dr["tipo_contenedor"].ToString());
                lista.Add(item);
            }

            cn.Close();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            Context.Response.Write(jss.Serialize(lista));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void eliminarRegTemporalSolIngCaja(long idTemporal)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand();
            cm.Connection = cn;
            cm.CommandText = "spDelRegistroTemporalSolicitudIngresoCaja @idTemporal=" + idTemporal.ToString();
            cm.ExecuteNonQuery();
            cn.Close();
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void delRegistroRecuperacionCaja_TMP(long idCaja)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand();
            cm.Connection = cn;
            cm.CommandText = "spdelRegistroRecuperacionCaja_TMP @idCaja=" + idCaja.ToString();
            cm.ExecuteNonQuery();
            cn.Close();
        }

        [WebMethod]
        public void consultacsv()
        {
            DataSet Ds = new DataSet();
            string respuesta = "";
            using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString))
            {
                SqlCommand cmd = new SqlCommand("select c.Nombre_Tabla, c.codigo, c.descripcion from SAF_Codificacion as c order by c.Nombre_Tabla, c.codigo", cnn);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(Ds);
            }
            respuesta=ConvertToCSV(Ds);
            Context.Response.Write(respuesta);
        }






        private string ConvertToCSV(DataSet objDataSet)
        {
            StringBuilder content = new StringBuilder();

            if (objDataSet.Tables.Count >= 1)
            {
                DataTable table = objDataSet.Tables[0];

                if (table.Rows.Count > 0)
                {
                    DataRow dr1 = (DataRow)table.Rows[0];
                    int intColumnCount = dr1.Table.Columns.Count;
                    int index = 1;

                    //add column names
                    foreach (DataColumn item in dr1.Table.Columns)
                    {
                        content.Append(String.Format("\"{0}\"", item.ColumnName));
                        if (index < intColumnCount)
                            content.Append(";");
                        else
                            content.Append("\r\n");
                        index++;
                    }

                    //add column data
                    foreach (DataRow currentRow in table.Rows)
                    {
                        string strRow = string.Empty;
                        for (int y = 0; y <= intColumnCount - 1; y++)
                        {
                            strRow += "\"" + currentRow[y].ToString() + "\"";

                            if (y < intColumnCount - 1 && y >= 0)
                                strRow += ";";
                        }
                        content.Append(strRow + "\r\n");
                    }
                }
            }

            byte[] bytes = Encoding.Default.GetBytes(content.ToString());
            return Encoding.Default.GetString(bytes);
        }




        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void verificarExistencia(string idServicio, string idFormato, string detalleJson)
        {

            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
            cn.Open();
            SqlCommand cm = new SqlCommand();
            cm.Connection = cn;
            SqlDataReader Dr;
            List<MensajeRetorno> lista = new List<MensajeRetorno>();
            MensajeRetorno mensaje = new MensajeRetorno();
            JavaScriptSerializer jss = new JavaScriptSerializer();
            DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));

            DataSet DsClaves = ServicioCamposClave(int.Parse(idServicio));
            DataTable DtClaves = DsClaves.Tables[0];

            DataSet DsCampos = campos_Formato_Servicio_DataTable(int.Parse(idFormato));
            DataTable DtCampos = DsCampos.Tables[0];

            string cadena;
            string campo;
            mensaje.mensaje = "0";
            mensaje.fecha = "";
            mensaje.hora = "";

            string clave = "";


            int numeroCaja = int.Parse( dt.Rows[0]["id_caja"].ToString());
            cm.CommandText = "spGetEstadoCaja @idCaja="+numeroCaja.ToString();
            Dr = cm.ExecuteReader();
            if (Dr.Read()) {
                mensaje.mensaje = "<p>El número de caja <b>ya esta registrado</b>, con el cliente " + Dr["nombreCliente"].ToString()+", estado actual: "+Dr["descripcionEstado"].ToString()+", fecha estado: "+Dr["fecha_estado"].ToString()+"</p>" ;
                lista.Add(mensaje); 
                Context.Response.Write(jss.Serialize(lista));
                Dr.Close();
                return;
            }
            Dr.Close();


            if (DtClaves.Rows.Count > 0)
            {
                for (int fl = 0; fl <= DtClaves.Rows.Count - 1; fl++)
                {

                    clave = DtClaves.Rows[fl][0].ToString();
                    for (int f2 = 0; f2 <= DtCampos.Rows.Count - 1; f2++)
                    {
                        if (DtClaves.Rows[fl][0].ToString() == DtCampos.Rows[f2][3].ToString()) //Busco el Campo Clave dentro de los Campos asignados al Formato
                        {
                            DtClaves.Rows[fl][1] = DtCampos.Rows[f2][0]; // El Orden (que indica la columna en el archivo)
                           // DtClaves.Rows[fl][2] = DtCampos.Rows[f2][4]; // El tipo del Campo
                        }
                    }
                }

                Boolean existe;
                cadena = "";

                for (int fl = 0; fl <= DtClaves.Rows.Count - 1; fl++)
                {

                    for (int cl = 0; cl <= dt.Columns.Count - 1; cl++)
                    {
                        if (DtClaves.Rows[fl][0].ToString() == dt.Columns[cl].ColumnName)
                        {

                            if (int.Parse(DtClaves.Rows[fl][2].ToString()) == 2)
                            {
                                campo = " and " + DtClaves.Rows[fl][0].ToString() + "=" + dt.Rows[0][cl].ToString();
                            }
                            else
                            {
                                campo = " and " + DtClaves.Rows[fl][0].ToString() + "='" + dt.Rows[0][cl].ToString() + "'";
                            }
                            cadena = cadena + campo;
                        }

                    }

                }


                cm.CommandText = "spGetVerificarExistenciaEnInventario @idCliente=" + idServicio.ToString() +", @filtro='" + cadena.Replace("'", "''")+"'";
                Dr = cm.ExecuteReader();
                existe = Dr.Read();
                Dr.Close();

                if (existe)
                {
                    mensaje.mensaje = "1";
                }

           }

           lista.Add(mensaje);
           Context.Response.Write(jss.Serialize(lista));

        }

    }
    
}
