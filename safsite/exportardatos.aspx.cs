using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Text;
using System.Web.Services;
using Newtonsoft.Json;
using System.Web.Script.Serialization;
using System.Web.Script.Services;
public partial class exportardatos : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string tipoInforme = "";
        string param1 = "";
        string param2 = "";
        string param3 = "";
        string param4 = "";
        string param5 = "";
        string param6 = "";

        tipoInforme = Request.QueryString["tipoInforme"];
        param1= Request.QueryString["param1"];
        param2 = Request.QueryString["param2"];
        param3= Request.QueryString["param3"];
        param4 = Request.QueryString["param4"];
        param5 = Request.QueryString["param5"];
        param6 = Request.QueryString["param6"];

        if ( tipoInforme=="1") {
            buscarCajas(param1, param2, param3);
        }
        if (tipoInforme == "2")
        {
            consultarRegistroPorTipoDocumento(param1, param2, param3);
        }
        if (tipoInforme == "3")
        {
            buscarSolicitudes(param1, param2, param3, param4, param5, param6);
        }

    }

    [WebMethod]
    public void buscarCajas(string idServicio, string desde, string hasta)
    {
        DataSet Ds = new DataSet();
         using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString))
        {
            string consulta= "spGetCajasParaCSV @idCliente = " + idServicio.ToString() + ", @desde = '" +desde.ToString() + "', @hasta = '" + hasta.ToString()+"'";
            SqlCommand cmd = new SqlCommand(consulta, cnn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(Ds);
        }
        ConvertToCSV(Ds);
        string strScript = "window.close();";
        ScriptManager.RegisterStartupScript(this, typeof(string), "key", strScript, true);
    }

    public void consultarRegistroPorTipoDocumento(string idServicio, string idFormato, string detalleJson)
    {
        JavaScriptSerializer jss = new JavaScriptSerializer();
        DataTable dt = (DataTable)JsonConvert.DeserializeObject(detalleJson, typeof(DataTable));
        // dt.Columns.RemoveAt(dt.Columns.Count - 1);
        string cadena = "id_cliente=" + idServicio + " and ";
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

        string campos = "select ";

        for (int x = 0; x <= dt1.Rows.Count - 1; x++)
        {
            campos = campos + "SAF_Inventario." + dt1.Rows[x]["campoNombreSafInventario"].ToString() + " as " + dt1.Rows[x]["CampoNombreCliente"].ToString().Replace(" ", "") +  ", ";
        }

        campos = campos + "Coalesce(SAF_Inventario.bodega,'') as Bodega, Coalesce(SAF_Inventario.bandeja,'') as Bandeja, SAF_Codificacion.Descripcion as Estado,  SUBSTRING(ltrim(Saf_inventario.Fecha_estado),7,2)+'/'+SUBSTRING(ltrim(Saf_inventario.Fecha_estado),5,2)+'/'+SUBSTRING(ltrim(Saf_inventario.Fecha_estado),1,4) as [Fecha Estado] , Coalesce(id_lote_solicitud,0) as [Nº Solicitud], coalesce(solicitante,'') as Solicitante";

        DataTable dt2 = new DataTable();
        cmd.Connection = cn;
        cadena = campos + " from SAF_Inventario inner join SAF_Codificacion on ( SAF_Codificacion.nombre_tabla='ESTADO_INV' and " +
            "SAF_Codificacion.Codigo=SAF_Inventario.Id_Estado) where " + cadena;

        cmd.CommandText = cadena;
        DataSet Ds = new DataSet();
        SqlDataAdapter d3 = new SqlDataAdapter(cmd);
        d3.Fill(Ds);
        cn.Close();
        ConvertToCSV(Ds);
        string strScript = "window.close();";
        ScriptManager.RegisterStartupScript(this, typeof(string), "key", strScript, true);
    }

    public void buscarSolicitudes(string idServicio, string tipoSolicitud, string estado, string desde, string hasta, string numeroSolicitud)
    {
        DataSet Ds = new DataSet();
        if (numeroSolicitud == "")
        {
            numeroSolicitud = "0";
        }

        if (desde != "")
        {
            desde = desde.Substring(6, 4) + desde.Substring(3, 2) + desde.Substring(0, 2);
        }
        if (hasta != "")
        {
            hasta = hasta.Substring(6, 4) + hasta.Substring(3, 2) + hasta.Substring(0, 2);
        }
        using (SqlConnection cnn = new SqlConnection(ConfigurationManager.ConnectionStrings["Dbconnection"].ConnectionString))
        {
            SqlCommand cmd = new SqlCommand("spGetSolicitudPorClienteServicioParaCSV @idCliente=" + idServicio + ", @idEstado=" + estado + ", @tipoSolicitud=" + tipoSolicitud +
                ", @fechaDesde='" + desde.ToString() + "', @fechaHasta='" + hasta.ToString() + "', @numeroSolicitud=" + numeroSolicitud, cnn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(Ds);
        }
        ConvertToCSV(Ds);
        string strScript = "window.close();";
        ScriptManager.RegisterStartupScript(this, typeof(string), "key", strScript, true);


    }

    private void ConvertToCSV(DataSet objDataSet)
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

        string respuesta = Encoding.Default.GetString(bytes);
        try
        {
            string csvFile = DateTime.Now.Ticks.ToString() + ".csv";
            string csvFilePath = Server.MapPath("~/archivos/") + csvFile;
            using (StreamWriter sw = new StreamWriter(@csvFilePath, false, System.Text.Encoding.UTF8))
            {
                sw.Write(respuesta.ToString());
                sw.Close();
            }

            //Preparamos el response ...
            Response.Clear();
            Response.AddHeader("Content-Disposition", "attachment; filename="+ csvFile);
            Response.ContentType = "application/vnd.csv";
            Response.Charset = "UTF-8";
            Response.ContentEncoding = System.Text.Encoding.UTF8;

            //Cargamos el archivo en memoria ...
            byte[] MyData = (byte[])System.IO.File.ReadAllBytes(csvFilePath);
            Response.BinaryWrite(MyData);

            //Eliminamos el archivo ...
            System.IO.File.Delete(csvFilePath);

            //Terminamos el response ..
            Response.End();

        }
        catch (Exception ex)
        {
            ClientScript.RegisterClientScriptBlock(this.GetType(), "CSV", "alert('Error al exportar el fichero')", true);
        }
    }


}