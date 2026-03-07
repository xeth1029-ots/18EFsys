<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_007.aspx.vb" Inherits="WDAIIP.SD_03_007" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員資料匯出</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript">
        function GETvalue() { document.getElementById('Button3').click(); }

        function SetOneOCID() { document.getElementById('Button4').click(); }

        function choose_class() {
            if (document.getElementById('OCID1').value == '') { document.getElementById('Button4').click(); }
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }

        function search() {
            var RedStar = document.getElementById('RedStar');
            if (RedStar && document.form1.OCIDValue1.value == '') {
                alert('請先選擇班別\n');
                return false;
            }
        }

        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;學員資料匯出</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>

                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Button4" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別 <font id="RedStar" runat="server" color="#ff0000">*</font> </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="display: none; left: 28%; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出欄位 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chkSort" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="4"></asp:CheckBoxList>
                                <input id="chkSortHidden" type="hidden" value="0" name="chkSortHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Button ID="Button1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button></td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
