<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_030.aspx.vb" Inherits="WDAIIP.SD_15_030" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>單位跨年度課程審查結果 </title>
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;單位跨年度課程審查結果 </asp:Label>
                </td>
            </tr>
        </table>
        <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
            <tr>
                <td class="bluecol_need">訓練機構 </td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" Width="70%" onfocus="this.blur()"></asp:TextBox>
                    <input id="Org" value="..." type="button" name="Org" runat="server" class="button_b_Mini" />
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                    <input id="ComidValue" type="hidden" name="ComidValue" runat="server" />
                      <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need" width="20%">年度區間</td>
                <td colspan="3" class="whitecol" width="80%">
                    <asp:DropDownList ID="yearlist1" runat="server"></asp:DropDownList>~
                    <asp:DropDownList ID="yearlist2" runat="server"></asp:DropDownList>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Button ID="bt_EXPORT1" runat="server" Text="匯出明細表" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="bt_EXPORT2" runat="server" Text="匯出總表" CssClass="asp_Export_M"></asp:Button>
                    <div align="center"></div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_YEAR_ROC1" runat="server" />
        <asp:HiddenField ID="Hid_YEAR_ROC2" runat="server" />
    </form>
</body>
</html>
