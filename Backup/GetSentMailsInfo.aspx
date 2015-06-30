<%@ Page Title="" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeFile="GetSentMailsInfo.aspx.cs" Inherits="GetSentMailsInfo" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="Server">
    <asp:ToolkitScriptManager ID="ToolkitScriptManager1" runat="server">
    </asp:ToolkitScriptManager>

    <table  id="tblLogin" align="center" width="60%" runat="server">
     <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
         <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
         <tr>
            <td>
                Enter Email Id:
            </td>
            <td>
                <asp:TextBox ID="txtUsrName" runat="server" ></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                Password:
            </td>
            <td>
                <asp:TextBox ID="txtpwd" runat="server" TextMode="Password"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr align="center">
            <td colspan="2">
                <asp:Button ID="btnSubmit" Text="Submit" runat="server" 
                    onclick="btnSubmit_Click" />
            </td>
        </tr>
    </table>


    <table id="tblEmails" align="center" width="80%" runat="server">
        <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr> <tr>
            <td>
                &nbsp;
            </td>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="6" align="center">
                <asp:Button ID="btnScanEmails" runat="server" Text="Scan Sent Emails" OnClick="btnScanEmails_onClick" />
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="6" align="center">
                <div id="divMsg" runat="server">
                </div>
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                From
            </td>
            <td>
                <%--<asp:Calendar ID="calFromDate" runat="server"></asp:Calendar>--%>
                <asp:CalendarExtender ID="calFromDate" runat="server" TargetControlID="txtFromDate"
                    PopupButtonID="imgBtnFromCalender">
                </asp:CalendarExtender>
                <asp:ImageButton runat="Server" ID="imgBtnFromCalender" ImageUrl="~/Images/Calendar_image.png"
                    AlternateText="Click here to display calendar" />
                <asp:TextBox ID="txtFromDate" runat="server"></asp:TextBox>
            </td>
            <td>
                &nbsp;
            </td>
            <td>
                To
            </td>
            <td>
                <%--<asp:Calendar ID="calToDate" runat="server"></asp:Calendar>--%>
                <asp:CalendarExtender ID="calToDate" runat="server" TargetControlID="txtToDate" PopupButtonID="imgBtnToCalender">
                </asp:CalendarExtender>
                <asp:ImageButton runat="Server" ID="imgBtnToCalender" ImageUrl="~/Images/Calendar_image.png"
                    AlternateText="Click here to display calendar" />
                <asp:TextBox ID="txtToDate" runat="server"></asp:TextBox>
            </td>
            <td>
                <asp:Button ID="btnExportToExcel" runat="server" Text="Export To Excel" OnClick="btnExportToExcel_onClick" />
            </td>
        </tr>        
        <tr>
            <td>
            </td>
            <td>
            </td>
        </tr>
    </table>

    </table>
          <tr>
            <td>
                <div id="divError" runat="server">
                </div>
            </td>
        </tr>
        <tr>
            <td>
                <div id="divEmail" runat="server">
                </div>
            </td>
        </tr>
    </table>
</asp:Content>
