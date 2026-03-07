<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_07_002.aspx.vb" Inherits="TIMS.SYS_07_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SYS_07_002</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
</head>
<body ms_positioning="FlowLayout">
    <form id="form1" method="post" runat="server">
    <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="600" border="0">
        <tbody>
            <tr>
                <td align="center">
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                首頁&gt;&gt;系統管理&gt;&gt;<font color="#990000">轉換失業週數代碼(99年度)</font>
                            </td>
                        </tr>
                    </table>
                    <asp:Button ID="bt_reserve" runat="server" Text="保留99年週數代碼" CssClass="asp_button_L"></asp:Button>&nbsp;
                    <asp:Button ID="bt_update" runat="server" Text="轉換99年失業週數代碼" CssClass="asp_button_L"></asp:Button>
                    <br />
                    <asp:Button ID="bt_restoreID" runat="server" Text="還原99年失業週數代碼" CssClass="asp_button_L"></asp:Button>&nbsp;
                    <asp:Button ID="bt_fixID" runat="server" Text="修正99年失業週數代碼" CssClass="asp_button_L"></asp:Button>
                </td>
            </tr>
        </tbody>
    </table>
    </form>
</body>
</html>
