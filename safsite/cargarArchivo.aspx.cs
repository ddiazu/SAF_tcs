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

public partial class cargarArchivo : System.Web.UI.Page
{
    public class campoServicio
    {
        public String nombre_campo;
        public String salida;
        public int tipo_valor;
        public int largo;
        public int orden;
        public string tabla_codificacion;

    }
    public string idUsuario;

    protected void Page_Load(object sender, EventArgs e)
    {
        tIdServicio.Text = Request.QueryString["idservicio"];
        tIdFormato.Text = Request.QueryString["idformato"];
        tNombreUsuario.Text = Request.QueryString["usuario"];
        idUsuario= Request.QueryString["idUsuario"];
        Mensaje.Text = "";
        Button1.Attributes.Add("onclick", " this.disabled = true; " + ClientScript.GetPostBackEventReference(Button1, null) + ";");

        this.FileUpload1.Attributes.Add("OnClick", "javascript:return limpiarMensaje();");
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
                long maximoTamanoBytes = (TamanoMaximoUploadMegaBytes * 1024)*1024;

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
                    lColumnas.Text = "";
                    lFilas.Text = "";
                    lColumnas.Text = "";
                    lErrores.Text = "";
                    RevisarInformacion(Server.MapPath("~/archivos/") + filename);
                }


            }
            catch (Exception ex)
            {
                Mensaje.Text = "Ha ocurrido un error al subir el archivo, vuelva a intentarlo.";
            }
        }
    }


    private void RevisarInformacion(string Archivo)
    {

        long idServicio = int.Parse(tIdServicio.Text);
        long idFormato = int.Parse(tIdFormato.Text);
        string usuario = tNombreUsuario.Text;
        long filas = 0;
        long columnas = 0;
        List1.Items.Clear();

        string linea;
        System.IO.StreamReader file = new System.IO.StreamReader(Archivo);
        linea = file.ReadLine();
        file.Close();

        //Verificar que en la cabecera no existan columnas numericas
        Char delimiter = ';';
        String[] substrings = linea.Split(delimiter);
        Boolean cabError = false;
        string tipo1;
        int numero;
        foreach (var substring in substrings)
        {
            if (int.TryParse(substring, out numero) | substring.Trim()=="")
            {
                cabError = true;
            }
        }

        if (cabError)
        {
            Mensaje.Text = "Existen errores en los nombres de las columnas en la primera fila (pueden haber nombres vacío o sólo números), corrija para subir el archivo.";
            return;
        }

        DataTable Dt = new DataTable();

         if (Archivo != "")
        {

            if (File.ReadAllLines(Archivo).Length< 2){
                Mensaje.Text = "El archivo debe contener una linea de cabecera y una o más línea con datos.";
                return;
            }

            Dt = ByteBufferToTable(File.ReadAllBytes(Archivo), true);
            filas = Dt.Rows.Count;
            columnas = Dt.Columns.Count;
        }

        lColumnas.Text = columnas.ToString();
        lFilas.Text = filas.ToString();

        DataSet Ds = new DataSet();
        DataSet DsClaves = new DataSet();
        safservicio.capaservicioSoapClient cliente =  new safservicio.capaservicioSoapClient();
        
        Ds = cliente.campos_Formato_Servicio_DataTable(idFormato);
        DsClaves = cliente.ServicioCamposClave(idServicio);

        DataTable DtClaves = new DataTable();
        DtClaves = DsClaves.Tables[0];



        int tipocampo;
        int tipovalor;
        int largo;
        string valor;
        string nombreCampo;
        List<string> errores = new List<string>();

        DataTable dt2 = new DataTable();

        dt2 = Ds.Tables[0];

        lColFormato.Text = (dt2.Rows.Count ).ToString();

        if (dt2.Rows.Count != columnas)
        {
            Mensaje.Text = "El archivo no puede ser procesado, el numero de columnas es distinto al formato definido.";
            return;
        }


        for (int r = 0; r <= dt2.Rows.Count - 1; r++)
        {


            tipocampo = int.Parse(dt2.Rows[r][6].ToString());
            largo = int.Parse(dt2.Rows[r][5].ToString());
            nombreCampo = dt2.Rows[r][3].ToString();

            bool isNum;
            double retNum;
            
            for (int x = 0; x <= Dt.Rows.Count - 1; x++)
            {
                valor = Dt.Rows[x][r].ToString();
                //Verificar Error de Tipo
                if (tipocampo == 2) //numerico
                 {
                    if (valor.Length > 0)
                    {
                        isNum = Double.TryParse(Convert.ToString(valor),
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.NumberFormatInfo.InvariantInfo,
                            out retNum);
                        if (!isNum)
                        {
                            errores.Add("Linea " + (x + 1).ToString() + ", Columna " + (r + 1).ToString() + ", Error de Tipo");
                        }
                    }
                    else
                    {
                        valor = "0";
                    }
                }
                if (tipocampo == 1) //texto
                {

                }
                if (tipocampo == 3)//fecha
                {

                }
                //Verificar Error de Largo
                if (largo > 0)
                {
                    if (largo < valor.Length && valor.Length > 0)
                    {
                        errores.Add("Linea " + (x + 1).ToString() + ", Columna " + (r + 1).ToString() + ", Error de Largo, Valor " + valor.ToString());
                    }
                }

            }

            
        }


        if (errores.Count>0)
        {
           
            foreach (var i in errores)
            {
                List1.Items.Add(i.ToString());
            }

            lErrores.Text = errores.Count.ToString();

            Mensaje.Text = "El archivo no puede ser procesado ya que contiene errores, corrijalos y vuelva a subirlo.";
            return;
        }

        //Verificar Existencia de Registro y Cual es su estado

        SqlConnection cn = new SqlConnection();
        cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
        cn.Open();
        SqlCommand cm = new SqlCommand();
        cm.Connection = cn;
        SqlDataReader Dr;
                
        if (DtClaves.Rows.Count > 0)
        {
            for (int fl = 0; fl <= DtClaves.Rows.Count - 1; fl++)
            {
                for (int f2 = 0; f2 <= dt2.Rows.Count - 1; f2++)
                {
                    if (DtClaves.Rows[fl][0].ToString() == dt2.Rows[f2][3].ToString()) //Busco el Campo Clave dentro de los Campos asignados al Formato
                    {
                        DtClaves.Rows[fl][2] = dt2.Rows[f2][0]; // El Orden (que indica la columna en el archivo)
                        DtClaves.Rows[fl][3] = dt2.Rows[f2][6]; // El tipo del Campo
                    }
                }
            }

            int posicion = 0;
            string campoIdCaja;
            for (int f3 = 0; f3 <= dt2.Rows.Count - 1; f3++)
            {
                campoIdCaja = dt2.Rows[f3][3].ToString().ToUpper() ;
                if ("ID_CAJA" == campoIdCaja) //Busco el Campo Clave dentro de los Campos asignados al Formato
                {
                    posicion = f3;
                }
            }

            string cadena;
            string campo;
            Boolean existe;
            int numCol;
            string nomCol;
            string numeroCaja;
            Boolean CajaExiste = false;
            for (int fl = 0; fl <= Dt.Rows.Count - 1; fl++)
            {

                numeroCaja = Dt.Rows[fl][posicion].ToString();

                cm.CommandText = "spGetEstadoCaja @idCaja=" + numeroCaja.ToString();
                Dr = cm.ExecuteReader();
                CajaExiste = Dr.Read();
                Dr.Close();

                if (!CajaExiste)
                {
                    cadena = "";
                    for (int f2 = 0; f2 <= DtClaves.Rows.Count - 1; f2++)
                    {
                        nomCol = DtClaves.Rows[f2][1].ToString();
                        numCol = 0;
                        for (int pos = 0; pos <= Dt.Columns.Count - 1; pos++)
                        {
                            if (Dt.Columns[pos].ColumnName == nomCol)
                            {
                                numCol = pos;
                            }

                        }
                        if (int.Parse(DtClaves.Rows[f2][3].ToString()) == 2)
                        {
                            campo = " and " + DtClaves.Rows[f2][0].ToString() + "=" + Dt.Rows[fl][numCol];
                        }
                        else
                        {
                            campo = " and " + DtClaves.Rows[f2][0].ToString() + "='" + Dt.Rows[fl][numCol] + "'";
                        }
                        cadena = cadena + campo;
                    }

                    cm.CommandText = "spGetVerificarExistenciaEnInventario @idCliente=" + idServicio.ToString() + ", @filtro='" + cadena.Replace("'", "''") + "'";
                    Dr = cm.ExecuteReader();
                    existe = Dr.Read();
                    Dr.Close();
                    if (existe)
                    {
                        errores.Add("Linea " + (fl + 1).ToString() + ", el registro que intenta subir como nuevo ya existe.");
                    }
                }
                else
                {
                    errores.Add("Linea " + (fl + 1).ToString() + ", el numero de caja que intenta subir ya existe.");
                }


               
            }
        }


        if (errores.Count > 0)
        {
            foreach (var i in errores)
            {
                List1.Items.Add(i.ToString());
            }

            lErrores.Text = errores.Count.ToString();

            Mensaje.Text = "El archivo no puede ser procesado ya que contiene errores, corrijalos y vuelva a subirlo.";
            return;
        }

        DataSet Ds2 = new DataSet();
        Ds2.Tables.Add(Dt);
        string respuesta;
        respuesta = "";
        respuesta= cliente.grabarSolicitudIngresoDataSet(idServicio.ToString(), idFormato.ToString(), usuario, idUsuario, Ds2);
           
        JavaScriptSerializer jss = new JavaScriptSerializer();
        DataTable dtresp = (DataTable)JsonConvert.DeserializeObject(respuesta, typeof(DataTable));
        if (dtresp.Rows.Count > 0)
        {
            if (int.Parse(dtresp.Rows[0][0].ToString()) > 0)
            {
                List1.Items.Clear();
                Mensaje.Text = "Se ha creado la Solicitud Número: " + dtresp.Rows[0][0].ToString();
                Button1.Enabled = true;
            }
        }
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
                            if (line != null){
                                lineArray = line.Split(delimiter);
                            }
                            
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


    public void grabarSolicitudIngresoDataSet(string idServicio, string idFormato, string Usuario, DataSet Ds)
    {

        DataTable dt = Ds.Tables[0];
        grabarSolicitudIngreso(idServicio, idFormato, Usuario, dt);
 
    }


    private void grabarSolicitudIngreso(string idServicio, string idFormato, string Usuario, DataTable dt)
    {
        var registros = 0;
        registros = dt.Rows.Count;


        List<campoServicio> campos = campos_Formato_Servicio(int.Parse(idFormato));


        if (campos.Count != dt.Columns.Count)
        {

            return ;
        }

        SqlConnection cn = new SqlConnection();
        cn.ConnectionString = ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString;
        cn.Open();
        SqlCommand cm = new SqlCommand("spInsCabeceraSolicitudIngreso @idServicio=" + idServicio + ", @idFormato=" + idFormato + ", @usuario='" + Usuario + "'", cn);
        SqlDataReader Dr = cm.ExecuteReader();
        long idLote = 0;
        long idUnidad = 0;
        long idTipoDoc = 0;
        if (Dr.Read())
        {
            idLote = int.Parse(Dr["IDLote"].ToString());
            idUnidad = int.Parse(Dr["idUnidad"].ToString());
            idTipoDoc = int.Parse(Dr["idTipoDoc"].ToString());
           }
        Dr.Close();

        string comando = "";
        string cadena = "Insert into SAF_OR_Detalle(ID_Lote_OR, id_cliente, id_unidad, id_unidad_cliente, id_estado, id_linea, id_TipoDoc, ";

        for (int c = 0; c <= campos.Count - 1; c++)
        {
            cadena += campos[c].nombre_campo + ", ";
        }
        cadena = cadena.Substring(0, cadena.Length - 2) + ") ";


        for (int r = 0; r <= dt.Rows.Count - 1; r++)
        {
            string valores = "Values(" + idLote.ToString() + ", " + idServicio + ", " + idUnidad.ToString() + ", 0, 1, " + (r + 1).ToString() + ", " + idTipoDoc.ToString() + ", ";
            for (int c = 0; c <= dt.Columns.Count - 1; c++)
            {
                if (campos[c].tipo_valor == 2) // Texto
                {
                    valores += "'" + dt.Rows[r][c].ToString() + "', ";
                }
                if (campos[c].tipo_valor ==1) // Numerico
                {
                    valores += dt.Rows[r][c].ToString() + ", ";
                }
                if (campos[c].tipo_valor == 3) // Fecha
                {
                    valores += "'" + dt.Rows[r][c].ToString() + "', ";
                }
            }
            comando = cadena + valores.Substring(0, valores.Length - 2) + ");";
            cm.CommandText = comando;
            cm.ExecuteNonQuery();
        }
        cn.Close();
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
            lista.Add(campo);
        }
        cn.Close();
        return (lista);
    }
}