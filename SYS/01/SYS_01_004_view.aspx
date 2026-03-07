<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_004_view.aspx.vb" Inherits="WDAIIP.SYS_01_004_view" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>已賦予使用者權限計畫</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table width="590" align="center" class="table_nw" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" style="width: 20%">年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddl_years" runat="server" AutoPostBack="True">
                    </asp:DropDownList>
                    <input id="hid_acc" type="hidden" name="hid_acc" runat="server">
                </td>
            </tr>
        </table>
        <table width="100%" align="center">
            <tr>
                <td style="height: 135px" align="center" valign="top">
                    <asp:DataGrid ID="DataGrid1" CssClass="font" runat="server" Visible="False" Width="100%" HeaderStyle-CssClass="TD_TD1" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="WhiteSmoke" />
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="AccName" HeaderText="姓名">
                                <ItemStyle HorizontalAlign="Center" Width="10%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlanName" HeaderText="計畫代碼">
                                <ItemStyle Width="45%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                <ItemStyle Width="45%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="RID" HeaderText="RID"></asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                    <asp:Label ID="msg" CssClass="font" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hidacc" runat="server" />
        <asp:HiddenField ID="Hidyears" runat="server" />
    </form>
</body>
</html>
