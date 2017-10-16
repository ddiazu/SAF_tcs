<%@ Page Language="C#" AutoEventWireup="true" CodeFile="cargarRecuperacionDocumentos.aspx.cs" Inherits="cargarRecuperacionDocumentos" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css" /> 
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/2.1.4/jquery.min.js"></script> 
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/js/bootstrap.min.js"></script>
    <style type="text/css">
        .auto-style1 {
            width: 97%;
        }
        .auto-style9 {
            height: 35px;
            text-align: right;
        }
        .auto-style16 {
            width: 78px;
            text-align: right;
            height: 35px;
        }
        .auto-style20 {
            padding: 15px;
            width: 530px;
        }
        .auto-style22 {
            width: 154px;
            height: 39px;
        }
        .auto-style23 {
            height: 39px;
        }
    </style>
    <script>
        function mostrar_procesar()
        {
            var file = document.getElementById('<% = FileUpload1.ClientID %>').value;
            var input = document.getElementById('<% = FileUpload1.ClientID %>');
            if (file == null || file == '') {
                alert('Seleccione el archivo a subir.');
                return false;
            }
            //DEFINE UN ARRAY CON LAS EXTENSIONES DE ARCHIVOS VALIDAS
            var extArray = new Array(".csv");
            //SE EXTRAE LA EXTENSION DEL ARCHIVO CON EL PUNTO INCLUIDP 
            var ext = file.slice(file.indexOf(".")).toLowerCase();
            var archivos = input.files;
            var tamano = archivos[0].size;
	        if ('.csv' != ext) {
		        alert("El archivo NO es un CSV");
                return false;
            }
	        var tamanoMaximo = (10 * 1024 * 1024);
	        if (tamano > tamanoMaximo) {
	            alert("El tamaño excede el máximo permitido(10Mb)");
	            return false;
	        }
            document.getElementById("procesando_div").style.display ="''";
            setTimeout("document.images['procesando_gif'].src='Content/imagenes/pleasewait.gif'", 200);
        }
    </script>
</head>

<body>

    <form id="form1" runat="server">
    <div class="container">
        <div>
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>    
            <asp:Label ID="tIdServicio" runat="server" Text="IdServicio" Visible="False"></asp:Label>
    
            <asp:Label ID="tIdFormato" runat="server" Text="IdFormato" Visible="False"></asp:Label>
    
            <asp:Label ID="tNombreUsuario" runat="server" Text="tNombreUsuario" Visible="False"></asp:Label>
            <br />
            <h4>Indique el archivo .CSV a cargar...</h4>

        </div>
        <asp:FileUpload ID="FileUpload1" runat="server" Width="400px" ccsClass="form-control"/>
        <br />
        <br />
        <div class="row">
            <div class="col-md-5">
                <p>Indique por cual Campo Clave realizará la solicitud:</p>
            </div>
            <div class="col-md-3">
                 <asp:DropDownList ID="cmbCampoClave" runat="server" Height="30px" Width="210px" CssClass="form-control" OnSelectedIndexChanged="cmbCampoClave_SelectedIndexChanged" ></asp:DropDownList>
            </div>
        </div>
        <br />
        <br />
        <asp:Button ID="Button1" runat="server" Text="Procesar el Archivo" Width="146px" CssClass="btn btn-info btn-xs" OnClick="Button1_Click" />   
        &nbsp;&nbsp;   
        <span id="procesando_div" style="display: none; position:absolute; text-align:center">
            <img id="procesando_gif" src="Content/imagenes/pleasewait.gif" />
        </span>
        <asp:Label ID="Mensaje" runat="server" Font-Bold="True" ForeColor="#0066CC" Text="Mensaje del Proceso"></asp:Label>
        <br />
        <br />
        <table class="nav-justified">
            <tr>
                <td class="auto-style22">Tipo de Envio</td>
                <td class="auto-style23">
        <asp:DropDownList ID="cmbTipoEnvio" runat="server" Height="30px" Width="210px" CssClass="form-control" >
        </asp:DropDownList>
                </td>
                <td class="auto-style23"></td>
                <td class="auto-style23"></td>
            </tr>
            <tr>
                <td class="auto-style22">Forma de Envio</td>
                <td class="auto-style23">
        <asp:DropDownList ID="cmbFormaEnvio" runat="server" Height="30px" Width="210px"  CssClass="form-control">
        </asp:DropDownList>
                </td>
                <td class="auto-style23"></td>
                <td class="auto-style23"></td>
            </tr>
            <tr>
                <td class="auto-style22">Tipo de Salida</td>
                <td class="auto-style23">
        <asp:DropDownList ID="cmbTipoSalida" runat="server" Height="30px" Width="210px"  CssClass="form-control">
        </asp:DropDownList>
                </td>
                <td class="auto-style23"></td>
                <td class="auto-style23"></td>
            </tr>
            <tr>
                <td class="auto-style22">Tiempo de Servicio</td>
                <td class="auto-style23">
        <asp:DropDownList ID="cmbTiempoServicio" runat="server" Height="30px" Width="210px"  CssClass="form-control">
        </asp:DropDownList>
                </td>
                <td class="auto-style23"></td>
                <td class="auto-style23"></td>
            </tr>
        </table>
        <asp:Button ID="btnEmitirSolicitud" runat="server" Text="Emitir Solicitud" Width="146px" CssClass="btn btn-success" OnClick="btnEmitirSolicitud_Click" />   
        <br />

        &nbsp;<br />
        <br />
            <div class="panel panel-primary">
              <div class="panel-heading">Resultados de la carga:</div>
              <div class="auto-style20">
                    <table class="auto-style1">
                        <tr>
                            <td style="width:150px">Filas</td>
                            <td class="auto-style16">
                                <asp:Label ID="lFilas" runat="server" BorderStyle="Solid" Text="0" Width="50px"></asp:Label>
                            </td>
                            <td style="width:140px">Errores Encontrados</td>
                            <td class="auto-style9">
                                <asp:Label ID="lErrores" runat="server" BorderStyle="Solid" Text="0" Width="50px"></asp:Label>
                            </td>
                        </tr>
                        </table>
                    <br />
                    <strong>Lista de Errores</strong><br />
                    <asp:GridView ID="Grid1" runat="server" BackColor="White" BorderColor="#CCCCCC" BorderStyle="Solid" BorderWidth="1px" CellPadding="3" EnableModelValidation="True">
                        <FooterStyle BackColor="White" ForeColor="#000066" />
                        <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
                        <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
                        <RowStyle ForeColor="#000066" />
                        <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
                    </asp:GridView>
                    <br />
              </div>
            </div>
        <br />
        <br />
        <br />
        <br />
        <br />
    </div>
    </form>
</body>
</html>