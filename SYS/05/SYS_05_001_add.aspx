<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_05_001_add.aspx.vb" Inherits="WDAIIP.SYS_05_001_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>上稿維護-FAQ_Detail</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;上稿維護&gt;&gt;FAQ</asp:Label>
                </td>
            </tr>
        </table>
        <div></div>
        <table class="table_nw" id="Table1" width="100%" cellspacing="1" cellpadding="1">
            <tbody>
                <tr>
                    <td class="bluecol" style="width: 20%">功能名稱
                    </td>
                    <td class="whitecol">
                        <asp:DropDownList ID="FunctionList" runat="server">
                        </asp:DropDownList>
                        <asp:RequiredFieldValidator ID="MustFunctionName" runat="server" ControlToValidate="FunctionList" Display="None" ErrorMessage="請選擇功能名稱"></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">提問單位
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="Textbox2" runat="server" Columns="40" Width="40%"></asp:TextBox><asp:RequiredFieldValidator ID="MustPostUnit" runat="server" ControlToValidate="Textbox2" Display="None" ErrorMessage="請輸入提問單位"></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">問題描述
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="TextBox3" runat="server" TextMode="MultiLine" Columns="50" Rows="6"></asp:TextBox><asp:RequiredFieldValidator ID="MustQuestion" runat="server" ControlToValidate="TextBox3" Display="None" ErrorMessage="請輸入問題描述"></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">解答描述
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="TextBox4" runat="server" TextMode="MultiLine" Columns="50" Rows="6"></asp:TextBox><asp:RequiredFieldValidator ID="MustDeal" runat="server" ControlToValidate="TextBox4" Display="None" ErrorMessage="請輸入解答描述"></asp:RequiredFieldValidator>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" class="whitecol" align="center">
                        <asp:Button ID="bt_addrow" Text="儲存" runat="server" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button1" runat="server" Text="回上頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </tbody>
        </table>
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" Width="280px" HeaderText="您尚未通過的橍位：" DisplayMode="List"></asp:ValidationSummary>
        <input id="AcceptSearch" type="hidden" name="AcceptSearch" runat="server">
    </form>
</body>
</html>
