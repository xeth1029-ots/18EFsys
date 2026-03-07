<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_03_004.aspx.vb" Inherits="TIMS.TC_03_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TC_03_004</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
</head>
<body ms_positioning="FlowLayout">
    <form id="form1" method="post" runat="server">
    <table class="font" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td class="font" colspan="4">
                首頁&gt;&gt;訓練機構管理&gt;&gt;<font color="#990000">班級申請匯入作業</font>
            </td>
        </tr>
        <tr class="table_sch">
            <td class="bluecol" width="100">
                匯入班級資料
            </td>
            <td colspan="3" class="whitecol">
                <input id="File1" type="file" name="File1" runat="server" size="50"><asp:Button ID="Btn_XlsImport"
                    runat="server" Text="匯入資料" CssClass="asp_button_M"></asp:Button><br />(必須為xls格式)
                <asp:HyperLink ID="HyperLink1" runat="server" ForeColor="#8080FF" NavigateUrl="../../Doc/Plan_Import28.zip"
                    CssClass="font">下載整批上載格式檔</asp:HyperLink>
            </td>
        </tr>
        <tr>
            <td colspan="4">
                <asp:DataGrid ID="dtgAddresses1" runat="server" CellPadding="4" CssClass="font">
                    <FooterStyle Wrap="False"></FooterStyle>
                    <SelectedItemStyle Wrap="False"></SelectedItemStyle>
                    <EditItemStyle Wrap="False"></EditItemStyle>
                    <AlternatingItemStyle Wrap="False" BackColor="#F5F5F5"></AlternatingItemStyle>
                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                    <PagerStyle Wrap="False"></PagerStyle>
                </asp:DataGrid>
            </td>
        </tr>
        <tr>
            <td colspan="4">
                <asp:DataGrid ID="dtgAddresses2" runat="server" CssClass="font">
                    <AlternatingItemStyle BackColor="#F5F5F5" />
                    <HeaderStyle CssClass="head_navy" />
                </asp:DataGrid>
            </td>
        </tr>
        <tr>
            <td colspan="4">
                <asp:DataGrid ID="dtgAddresses3" runat="server" CssClass="font">
                    <AlternatingItemStyle BackColor="#F5F5F5" />
                    <HeaderStyle CssClass="head_navy" />
                </asp:DataGrid>
            </td>
        </tr>
        <tr>
            <td colspan="4">
                <asp:DataGrid ID="dtgAddresses4" runat="server" CssClass="font">
                    <AlternatingItemStyle BackColor="#F5F5F5" />
                    <HeaderStyle CssClass="head_navy" />
                </asp:DataGrid>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
