<%@ Page Language="C#" AutoEventWireup="true" CodeFile="cargarArchivo.aspx.cs" Inherits="cargarArchivo" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <link href="Content/css/bootstrap.min.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .auto-style1 {
            width: 97%;
        }
        .auto-style3 {
            width: 120px;
            height: 35px;
        }
        .auto-style7 {
            width: 24px;
            text-align: right;
            height: 36px;
        }
        .auto-style9 {
            width: 24px;
            height: 35px;
            text-align: right;
        }
        .auto-style10 {
            width: 120px;
            font-size: smaller;
            height: 36px;
        }
        .auto-style13 {
            display: block;
            padding: 6px 12px;
            font-size: 14px;
            line-height: 1.42857143;
            color: #555;
            background-color: #fff;
            background-image: none;
            border: 1px solid #ccc;
            border-radius: 4px;
            -webkit-box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
            box-shadow: inset 0 1px 1px rgba(0,0,0,.075);
            -webkit-transition: border-color ease-in-out .15s,-webkit-box-shadow ease-in-out .15s;
            -o-transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
            transition: border-color ease-in-out .15s,box-shadow ease-in-out .15s;
        }
        .auto-style15 {
            width: 139px;
            font-size: smaller;
            height: 35px;
        }
        .auto-style16 {
            width: 78px;
            text-align: right;
            height: 35px;
        }
        .auto-style18 {
            width: 139px;
            font-size: small;
            height: 36px;
        }
        .auto-style19 {
            width: 78px;
            text-align: right;
            height: 36px;
        }
        .auto-style20 {
            padding: 15px;
            width: 537px;
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
        function limpiarMensaje() {
            document.getElementById("Mensaje").innerHTML = "";
            document.getElementById("lFilas").innerHTML = "";
            document.getElementById("lColumnas").innerHTML = "";
            document.getElementById("lColFormato").innerHTML = "";
            document.getElementById("lErrores").innerHTML = "";
            var listBox = document.getElementById("List1");
            listBox.options.length = 0;
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
        <asp:Button ID="Button1" runat="server" Text="Procesar el Archivo" Width="146px" OnClick="Button1_Click" CssClass="btn btn-info btn-xs" OnClientClick="return mostrar_procesar();" />   
        &nbsp;&nbsp;   
        <span id="procesando_div" style="display: none; position:absolute; text-align:center">
            <img id="procesando_gif" src="Content/imagenes/pleasewait.gif" />
        </span>
        <asp:Label ID="Mensaje" runat="server" Font-Bold="True" ForeColor="#0066CC" Text="Mensaje del Proceso"></asp:Label>
        <br />
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
                            <td style="width:140px">Columnas</td>
                            <td class="auto-style9">
                                <asp:Label ID="lColumnas" runat="server" BorderStyle="Solid" Text="0" Width="50px"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td style="width:150px">Columnas del Formato</td>
                            <td class="auto-style19">
                                <asp:Label ID="lColFormato" runat="server" BorderStyle="Solid" Text="0" Width="50px"></asp:Label>
                            </td>
                            <td style="width:140px">Errores Encontrados</td>
                            <td class="auto-style7">
                                <asp:Label ID="lErrores" runat="server" BorderStyle="Solid" Text="0" Width="50px"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <br />
                    <strong>Lista de Errores</strong><br />
                    <asp:ListBox ID="List1" runat="server" Height="127px" Width="525px" CssClass="auto-style13"></asp:ListBox>
              </div>
            </div>
        <br />
        <br />
        <br />
        <br />
        <br />
    </div>
    </form>
    <script src="Scripts/jquery.min.js"></script>
    <script src="Scripts/bootstrap.min.js"></script>
</body>
</html>