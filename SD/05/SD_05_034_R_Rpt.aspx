<%@ Page Language="VB" AutoEventWireup="false" Inherits="WDAIIP.SD_05_034_R_Rpt" CodeBehind="SD_05_034_R_Rpt.aspx.vb" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
</head>
<body>
    <%--<object style="display: none" id="factory" codebase="../../scriptx/ScriptX.cab#Version=6,2,433,14" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
    <form id="form1" runat="server">
    <asp:HiddenField ID="hidOCID" runat="server" />
    <asp:HiddenField ID="hidTMID" runat="server" />
    <asp:HiddenField ID="hidTPlanID" runat="server" />
    <asp:HiddenField ID="hidRID" runat="server" />
    <asp:HiddenField ID="hidSDate" runat="server" />
    <asp:HiddenField ID="hidEDate" runat="server" />
    <asp:HiddenField ID="hidItem1" runat="server" />
    <asp:HiddenField ID="hidItem2" runat="server" />
    <asp:HiddenField ID="hidItem3" runat="server" />
    <asp:HiddenField ID="hidItem4" runat="server" />
    <asp:HiddenField ID="hidUserID" runat="server" />
    <asp:HiddenField ID="hid" runat="server" />
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr id="trBtn" runat="server">
            <td align="right">
                <asp:Button ID="btnExport" runat="server" Text="匯出明細" CssClass="asp_Export_M" />
                <asp:Button ID="btnPrt" runat="server" Text="列印" CssClass="asp_Export_M" />
                <asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="asp_button_S" />
            </td>
        </tr>
        <tr>
            <td>
                <div id="div_print" class="Section1" runat="server">
                </div>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
