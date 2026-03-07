<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_008.aspx.vb" Inherits="WDAIIP.SD_02_008" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>學員報名甄試結果匯出</title>
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
        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }

        function search() {
            if (document.form1.OCIDValue1.value == '') {
                alert('請先選擇班別\n');
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
                    <asp:Label ID="titlelab1" runat="server"></asp:Label>
                    <%--<asp:Label ID="titlelab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;甄試及錄取&gt;&gt;學員報名甄試結果匯出</asp:Label>--%>
                    <asp:Label ID="titlelab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;學員報名甄試結果匯出</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3" CssClass="asp_button_M"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="Trwork2013a" runat="server">
                            <td class="bluecol">就服單位協助報名</td>
                            <td class="whitecol">
                                <asp:RadioButtonList Style="z-index: 0" ID="rblEnterPathW" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出欄位</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="sort" runat="server" RepeatColumns="5" RepeatLayout="flow" RepeatDirection="horizontal" CssClass="font">
                                </asp:CheckBoxList>
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
                                <asp:Button ID="button1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button></td>
                        </tr>
                    </table>
                    <div align="center">&nbsp;<asp:Label ID="msg" runat="server" ForeColor="red" CssClass="font"></asp:Label></div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
