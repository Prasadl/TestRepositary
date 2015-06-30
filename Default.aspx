<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeFile="Default.aspx.cs" Inherits="_Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
   <table align="center">
    <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>
 <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>

    <tr>
        <td> Enter User Name: </td>
        <td><asp:TextBox ID="txtUsrName" runat="server" TextMode="Password"></asp:TextBox></td>
    </tr>
    <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>

     <tr>
        <td> Password: </td>
        <td><asp:TextBox ID="txtpwd" runat="server" TextMode="Password"></asp:TextBox></td>
    </tr>

    <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
    </tr>

     <tr align="center" >
        <td colspan="2"><asp:Button ID="btnSubmit" Text="Submit" runat="server" /></td>
        
    </tr>


   </table>
</asp:Content>
