using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Globalization;
using System.Data.SqlClient;
using System.Data;
using System.Web.Script.Serialization;
using System.Web.Script.Services;
using Newtonsoft.Json;
using System.Configuration;
using System.Net.Mail;

public partial class cargarRecuperacionCajas : System.Web.UI.Page
{
    public string idUsuario;
   
    public class Caja
    {
        public long id_registro { get; set; }
        public int linea { get; set; }
        public string nro_caja { get; set; }
        public string estado { get; set; }
    }

    public class tipoEnvio
    {
        public string codigo { get; set; }
        public string descripcion { get; set; }

        public override string ToString()
        {
            return descripcion;
        }
    }

    public class tiempoRespuesta
    {
        public string codigo { get; set; }
        public string descripcion { get; set; }
        public override string ToString()
        {
            return descripcion;
        }
    }
    public class tipoSalida
    {
        public string codigo { get; set; }
        public string descripcion { get; set; }
        public override string ToString()
        {
            return descripcion;
        }
    }

    public class formaEnvio
    {
        public string codigo { get; set; }
        public string descripcion { get; set; }
        public override string ToString()
        {
            return descripcion;
        }
    }


    protected void Page_Load(object sender, EventArgs e)
    {
         tIdServicio.Text = Request.QueryString["idservicio"];
         tNombreUsuario.Text = Request.QueryString["usuario"];
         idUsuario = Request.QueryString["idUsuario"];

        Mensaje.Text = "";
        btnEmitirSolicitud.Enabled = false;
        Button1.Attributes.Add("onclick", " this.disabled = true; " + ClientScript.GetPostBackEventReference(Button1, null) + ";");
        SqlConnection cn = new SqlConnection();
        cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
        cn.Open();
        SqlCommand cm = new SqlCommand();
        cm.Connection = cn;
        cm.CommandText = "spGetTipoEnvioClienteServicio @idcliente=" + tIdServicio.Text;

        SqlDataReader Dr = cm.ExecuteReader();
        List<tipoEnvio> lsTipoEnvio = new List<tipoEnvio>();
        while (Dr.Read()) {
            tipoEnvio tp = new tipoEnvio();
            tp.codigo= Dr["codigo"].ToString();
            tp.descripcion = Dr["descripcion"].ToString();
            lsTipoEnvio.Add(tp);
        }
        Dr.Close();

        cm.CommandText = "spGetTiempoRespuestaClienteServicio @idcliente=" + tIdServicio.Text;
        Dr = cm.ExecuteReader();
        List<tiempoRespuesta> lsTiempoRespuesta = new List<tiempoRespuesta>();
        while (Dr.Read())
        {
            tiempoRespuesta tr = new tiempoRespuesta();
            tr.codigo = Dr["codigo"].ToString();
            tr.descripcion = Dr["descripcion"].ToString();
            lsTiempoRespuesta.Add(tr);
        }
        Dr.Close();

        cm.CommandText = "select codigo, descripcion from SAF_Codificacion where nombre_tabla = 'FORMA_ENVIO' order by codigo";
        Dr = cm.ExecuteReader();
        List<formaEnvio> lsFormaEnvio = new List<formaEnvio>();
        while (Dr.Read())
        {
            formaEnvio fe = new formaEnvio();
            fe.codigo = Dr["codigo"].ToString();
            fe.descripcion = Dr["descripcion"].ToString();
            lsFormaEnvio.Add(fe);
        }
        Dr.Close();

        cm.CommandText = "select codigo, descripcion from SAF_Codificacion where nombre_tabla='TIPO_SALIDA' order by codigo";
        Dr = cm.ExecuteReader();
        List<tipoSalida> lsTipoSalida = new List<tipoSalida>();
        while (Dr.Read())
        {
            tipoSalida ts = new tipoSalida();
            ts.codigo = Dr["codigo"].ToString();
            ts.descripcion = Dr["descripcion"].ToString();
            lsTipoSalida.Add(ts);
        }
        Dr.Close();

        cmbFormaEnvio.DataValueField = "codigo";
        cmbFormaEnvio.DataTextField = "descripcion";
        cmbFormaEnvio.DataSource = lsFormaEnvio.ToList();
        cmbFormaEnvio.DataBind();

        cmbTiempoServicio.DataValueField = "codigo";
        cmbTiempoServicio.DataTextField = "descripcion";
        cmbTiempoServicio.DataSource = lsTiempoRespuesta.ToList();
        cmbTiempoServicio.DataBind();

        cmbTipoEnvio.DataValueField = "codigo";
        cmbTipoEnvio.DataTextField = "descripcion";
        cmbTipoEnvio.DataSource = lsTipoEnvio.ToList();
        cmbTipoEnvio.DataBind();

        cmbTipoSalida.DataValueField = "codigo";
        cmbTipoSalida.DataTextField = "descripcion";
        cmbTipoSalida.DataSource = lsTipoSalida.ToList();
        cmbTipoSalida.DataBind();

        habilitarEnvio(false);
    }

    protected void habilitarEnvio(Boolean habilitar)
    {
        cmbFormaEnvio.Enabled = habilitar;
        cmbTiempoServicio.Enabled = habilitar;
        cmbTipoEnvio.Enabled = habilitar;
        cmbTipoSalida.Enabled = habilitar;
        btnEmitirSolicitud.Enabled = habilitar;
    }
    protected void Button1_Click(object sender, EventArgs e)
    {
        if (FileUpload1.HasFile)
        {
            try
            {

                string fileName = Server.HtmlEncode(FileUpload1.FileName);

                // Get the extension of the uploaded file.
                string extension = System.IO.Path.GetExtension(fileName);
                int TamanoMaximoUploadMegaBytes = int.Parse(ConfigurationManager.AppSettings.Get("TamanoMaximoUploadMegaBytes"));
                //Conversion de Mb y Bytes
                long maximoTamanoBytes = (TamanoMaximoUploadMegaBytes * 1024) * 1024;

                if (FileUpload1.FileContent.Length > maximoTamanoBytes)
                {
                    Mensaje.Text = "El archivo excede el tamaño máximo (" + TamanoMaximoUploadMegaBytes.ToString() + "MB)";
                    return;
                }
                // Allow only files with .doc or .xls extensions
                // to be uploaded.
                if ((extension != ".csv"))
                {
                    Mensaje.Text = "Solo se permiten archivos de tipo CSV";
                    Button1.Enabled = true;
                    return;
                }

                else
                {
                    string filename = Path.GetFileName(FileUpload1.FileName);
                    FileUpload1.SaveAs(Server.MapPath("~/archivos/") + filename);
                    Mensaje.Text = "El archivo ha subido correctamente para ser procesado.";
                    lFilas.Text = "";
                    lErrores.Text = "";
                    RevisarInformacion(Server.MapPath("~/archivos/") + filename);
                }


            }
            catch (Exception ex)
            {
                Mensaje.Text = "Ha ocurrido un error al subir el archivo, vuelva a intentarlo nuevamente.";
            }
        }
    }

    private void RevisarInformacion(string Archivo)
    {
        long idServicio = int.Parse(tIdServicio.Text);
        string usuario = tNombreUsuario.Text;
        long filas = 0;
        long columnas = 0;

        DataTable Dt = new DataTable();
        if (Archivo != "")
        {
            Dt = ByteBufferToTable(File.ReadAllBytes(Archivo), true);
            filas = Dt.Rows.Count;
            columnas = Dt.Columns.Count;
        }

        lFilas.Text = filas.ToString();

        SqlConnection cn = new SqlConnection();
        cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
        cn.Open();
        SqlCommand cm = new SqlCommand();
        cm.Connection = cn;

        string cadena = "spDelRecuperaCajasTemporal @idcliente=" + idServicio.ToString()+", @idusuario="+idUsuario.ToString()+";";
        cm.CommandText = cadena;
        cm.ExecuteNonQuery();
        for (int r = 0; r <= Dt.Rows.Count - 1; r++)
        {
            cadena = "spInsTempRecuperaCaja @idCliente=" +
                idServicio.ToString() + ", @idUsuario="+idUsuario.ToString() + ", @idCaja=" + Dt.Rows[r][0].ToString() +", @filaCarga=" + (r+1).ToString() + ";";
            cm.CommandText = cadena;
            cm.ExecuteNonQuery();

        }

        cm.CommandText = "spGetVerificarCajasSolicitadas @idCliente=" + idServicio.ToString()+ ", @idUsuario=" + idUsuario.ToString();
        SqlDataReader Dr = cm.ExecuteReader();
        List<Caja> lista = new List<Caja>();
        while (Dr.Read()){
            if (Dr["ErroresEncontrados"].ToString().Trim() !="")
            {
                Caja caja = new Caja();
                caja.nro_caja = Dr["id_Caja"].ToString();
                caja.linea = int.Parse(Dr["filaCarga"].ToString());
                caja.estado = Dr["ErroresEncontrados"].ToString();
                lista.Add(caja);
            }
        }
        Dr.Close();
        cn.Close();
        Grid1.AutoGenerateColumns = true;
        Grid1.DataSource = lista.ToList();
        Grid1.DataBind();
   //     Grid1.Columns[0].Visible = false;
        lErrores.Text = lista.Count.ToString();
        if (lista.Count > 0)
        {
            habilitarEnvio(false);
            Mensaje.Text = "Se encontraron errores en el archivo, corríjalos y vuelva a subirlo.";
        }else
        {
            //procesar
            habilitarEnvio(true);
            Mensaje.Text = "";

        }
        cn.Close();

    }

    public static DataTable ByteBufferToTable(byte[] buffer, bool includeHeader)
    {
        DataTable result = new DataTable();

        // Se asume que el separador de decimales es punto "." y el de miles "," (aunque este ultimo no se usa) 
        CultureInfo culture = System.Threading.Thread.CurrentThread.CurrentCulture;
        System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
        Dictionary<string, int> indexs = new Dictionary<string, int>();
        DataTable dt = new DataTable();
        char[] delimiter = new char[] { ';' };
        char[] zero = new char[] { '0' };

        using (StreamReader sr = new StreamReader(new MemoryStream(buffer)))
        {
            try
            {
                int rowsCompleted = 0;
                int lastLength = 0;
                bool readHeader = true;

                while (sr.Peek() > -1)
                {
                    bool addLine = true;
                    string line = sr.ReadLine();
                    string[] lineArray = line.Split(delimiter);

                    //Se chequea que tanto el orden como el nombre de las columnas correspondan según el orden dado
                    if (readHeader)
                    {
                        if (includeHeader)
                        {
                            int j = 0;
                            foreach (string column in lineArray)
                            {
                                DataColumn c = new DataColumn(column);
                                dt.Columns.Add(c);
                                indexs.Add(column, j);
                                j++;
                            }
                            //Se continua con la lectura del archivo
                            line = sr.ReadLine();
                            lineArray = line.Split(delimiter);
                        }
                        else
                        {
                            //Agrego columnas con nombre estandar, no se pasa a la siguiente linea del doc
                            for (int j = 0; j < lineArray.Length; j++)
                            {
                                DataColumn c = new DataColumn("Column" + j);
                                dt.Columns.Add(c);
                                indexs.Add("Column" + j, j);
                            }
                        }
                        //Se cambia el estado de esta variable para no volver a chequear el header
                        readHeader = false;
                    }

                    DataRow nuevaFila = dt.NewRow();
                    if (lastLength > 0 && lastLength != lineArray.Length)
                    {
                        continue;
                    }
                    lastLength = lineArray.Length;
                    try
                    {
                        foreach (DataColumn column in dt.Columns)
                        {
                            int index = indexs[column.ColumnName];
                            string colName = column.ColumnName;
                            string value = lineArray[index];
                            nuevaFila[colName] = string.IsNullOrEmpty(value) ? DBNull.Value + "" : value;
                        }
                    }
                    catch (Exception e)
                    {
                        throw e;
                    }
                    if (addLine)
                    {
                        dt.Rows.Add(nuevaFila);
                        rowsCompleted++;
                    }
                }
                return dt;
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = culture;
            }
        }
    }

    protected void btnSolicitud_click(object sender, EventArgs e)
    {

        string cadena = "spInsCabaceraSolicitudRecuperacionCajas @idServicio=" + tIdServicio.Text + ", @idUsuario=" + idUsuario.ToString()
            + ", @observaciones='', @tipoSolicitud=2, @tipoEnvio=" + cmbTipoEnvio.SelectedValue.ToString() + ", @formaEnvio=" + cmbFormaEnvio.SelectedValue.ToString()
            + ", @tipoSalida=" + cmbTipoSalida.SelectedValue.ToString() + ", @tiempoServicio=" + cmbTiempoServicio.SelectedValue.ToString() + 
            ", @nombreUsuario='"+ tNombreUsuario.Text + "';";


        SqlConnection cn = new SqlConnection();
        cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
        cn.Open();
        SqlCommand cm = new SqlCommand();
        cm.Connection = cn;
        cm.CommandText = cadena;
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
        }
        Dr.Close();

        cadena = "spInsCargaMasivaCajas @idLote=" + idSolicitud.ToString() + ", @idCliente=" + tIdServicio.Text + ", @idUsuario=" +
            idUsuario.ToString() + ", @idUnidad=" + idUnidad.ToString() + ", @nombreUsuario='', @idTipoDoc=" + idTipoDoc.ToString();

        cm.CommandText = cadena;

        cm.ExecuteNonQuery();

        cn.Close();

        Mensaje.Text = "Se ha creado la solicitud de retiro de cajas Numero: " + idSolicitud.ToString();

        btnEmitirSolicitud.Enabled = false;
        lErrores.Text = "";
        lFilas.Text = "";
        EnviarCorreoSolicitudRecuperacionCajas(idSolicitud, nombreCliente, "", "2");
    }

    private void EnviarCorreoSolicitudRecuperacionCajas(long idLote, string nombreCliente, string nombreUsuario, string tipoSolicitud)
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

}