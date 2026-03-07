<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_010_R.aspx.vb" Inherits="WDAIIP.SD_02_010_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>繳費通知單內容</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function choose_class() { openClass('SD_02_ch.aspx'); }

        function ReportPrint() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇班別!\n';
            if (!isChecked(document.form1.SelResult)) msg += '請選擇通知對象!\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }

        function checkTextLength(obj, long) {  //限定textbox的欄位長度
            var maxlength = new Number(long);
            if (obj.value.length > maxlength) {
                obj.value = obj.value.substring(0, maxlength);
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="titlelab1" runat="server"></asp:Label>
                    <%--<asp:Label ID="titlelab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;甄試及錄取&gt;&gt;繳費通知單內容</asp:Label>--%>
                    <asp:Label ID="titlelab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;繳費通知單內容</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol_need" style="width: 20%">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="hidden1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="hidden2" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="historytable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">通知對象</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="SelResult" runat="server" CssClass="font" RepeatDirection="horizontal"></asp:RadioButtonList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">郵寄類別</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="Mailtype1" runat="server" RepeatDirection="horizontal" CssClass="font">
                                    <asp:ListItem Value="1">印刷品</asp:ListItem>
                                    <asp:ListItem Value="2">平信</asp:ListItem>
                                    <asp:ListItem Value="3">限時</asp:ListItem>
                                    <asp:ListItem Value="4">掛號</asp:ListItem>
                                    <asp:ListItem Value="5">雙掛號</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Button ID="Button1" runat="server" Text="列印" Enabled="false" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="Button2" runat="server" Text="設定通知單內容" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="table11" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td style="background: #f1f9fc" align="left">正取繳費內容</td>
                        </tr>
                        <tr>
                            <td>
                                <div align="left">
                                    <strong></strong>
                                    <asp:TextBox ID="ItemVar_1" onblur="checkTextLength(this,512)" onkeyup="checkTextLength(this,512)" Width="100%" Rows="3" Columns="80" TextMode="multiline" runat="server" Height="89px" onchange="checkTextLength(this,512)"></asp:TextBox>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td style="background: #f1f9fc" align="left">備取繳費內容</td>
                        </tr>
                        <tr>
                            <td>
                                <div align="left">
                                    <strong></strong>
                                    <asp:TextBox ID="itemvar_2" Width="100%" Rows="3" Columns="80" TextMode="MultiLine" runat="server" Height="89px" onkeyup="checkTextLength(this,512)" onchange="checkTextLength(this,512)" onblur="checkTextLength(this,512)"></asp:TextBox>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button5" runat="server" Text="儲存" Visible="false" CssClass="asp_button_M"></asp:Button>
                        <br />
                        <asp:Label class="font" ID="msg" runat="server" Visible="false" ForeColor="red">尚未設定通知單內容</asp:Label>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
