<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="LerExcel.aspx.cs" Inherits="LerExcel.LerExcel" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
    <head runat="server">
        <title>Ler Excel com C#</title>
        <link type="text/css" rel="Stylesheet" href="Scripts/Planilha.css" />
    </head>
    <body>
        <form id="fLerExcel" runat="server">
            <asp:Label ID="lbArquivo" runat="server" Text="Selecione a planilha para leitura"></asp:Label>
            <br />
            <asp:FileUpload ID="fuArquivo" runat="server" />
            <asp:Button ID="btnLer" runat="server" Text="Ler Excel" style="float:right" onclick="btnLer_Click" />
            <hr />
        
            <div id="Tabela">
                <asp:Literal ID="lTabela" runat="server"></asp:Literal>
            </div>
        </form>
    </body>
</html>      
