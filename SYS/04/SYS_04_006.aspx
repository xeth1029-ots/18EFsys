<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_006.aspx.vb" Inherits="WDAIIP.SYS_04_006" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>天數設定</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        function check_data() {
            var msg = '';
            if (document.form1.Days1.value == '') msg += '請輸入非系統管理者限制天數\n';
            else if (!isUnsignedInt(document.form1.Days1.value)) msg += '非系統管理者限制天數必須為數字\n';
            if (document.form1.Days2.value == '') msg += '請輸入系統管理者限制天數\n';
            else if (!isUnsignedInt(document.form1.Days2.value)) msg += '系統管理者限制天數必須為數字\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;天數設定</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;<font color="#990000">天數設定</font>
						</td>
					</tr>
				</table>--%>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">非系統管理者限制天數
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="Days1" runat="server" Columns="5" Width="8%"></asp:TextBox>
                                天
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">系統管理者限制天數
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="Days2" runat="server" Columns="5" Width="8%"></asp:TextBox>天
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
