<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_006_edit.aspx.vb" Inherits="WDAIIP.TC_01_006_edit" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TC_01_006_edit</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        function chk(nn, text) {
            if (document.form1.TextBox2.value != '') {
                var num = '0123456789';
                var cmpChar;
                var chknum = nn.value

                for (var i = 0; i <= chknum.length - 1; i++) {
                    cmpChar = chknum.substring(i, i + 1);

                    if (num.indexOf(cmpChar) < 0) {
                        window.alert(text + ' 必須輸入數字!');
                        nn.focus();
                        return false;
                    }
                }
            }
        }

        function chkdata() {
            var msg = '';

            if (document.form1.DropDownList1.selectedIndex == '0') {
                msg = "必須選擇內外聘\n";
            }
            if (document.form1.DropDownList2.selectedIndex == '0') {
                msg = msg + "必須選擇師資別種類\n";
            }
            if (document.form1.TextBox1.value == '') {
                msg = msg + "必須填寫師資別名稱\n";
            }
            if (document.form1.TextBox2.value == '') {
                msg = msg + "必須填寫基本時數\n";
            }
            if (document.form1.TextBox3.value == '') {
                msg = msg + "必須填寫最高請領時數";
            }
            if (msg != '') {
                window.alert(msg);
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;師資別設定</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="Table3" width="100%" cellpadding="1" cellspacing="1">
                        <tr>
                            <td width="20%" class="bluecol_need">內外聘</td>
                            <td class="whitecol" width="30%">
                                <asp:DropDownList ID="DropDownList1" runat="server">
                                    <asp:ListItem Value="0">請選擇</asp:ListItem>
                                    <asp:ListItem Value="1">內</asp:ListItem>
                                    <asp:ListItem Value="2">外</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                            <td width="20%" class="bluecol_need">師資別種類</td>
                            <td class="whitecol" width="30%">
                                <asp:DropDownList ID="DropDownList2" runat="server">
                                    <asp:ListItem Value="0">請選擇</asp:ListItem>
                                    <asp:ListItem Value="1">訓練師類</asp:ListItem>
                                    <asp:ListItem Value="2">行政人員類</asp:ListItem>
                                    <asp:ListItem Value="3">外聘</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">師資別名稱</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TextBox1" runat="server" Width="35%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">基本時數</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TextBox2" runat="server" Width="20%"></asp:TextBox></td>
                            <td class="bluecol_need">最高請領時數</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TextBox3" runat="server" Width="20%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">ㄧ般鐘點費(學科)</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TextBox4" runat="server" Width="8%"></asp:TextBox>/時，(外聘鐘點費)</td>
                        </tr>
                        <tr>
                            <td class="bluecol">ㄧ般鐘點費(術科)</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TextBox5" runat="server" Width="8%"></asp:TextBox>/時</td>
                        </tr>
                        <tr>
                            <td class="bluecol">超時鐘點費</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TextBox6" runat="server" Width="8%"></asp:TextBox>/時</td>
                        </tr>
                        <tr>
                            <td class="bluecol">停止使用&nbsp;</td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBox ID="CB_NOUSE" runat="server" />（勾選停止使用）&nbsp;</td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;
                                <asp:Button ID="btnBack1" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button>
                                <%--<input id="Button2" type="button" value="回上一頁" name="Button2" runat="server" class="button_b_S">--%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
