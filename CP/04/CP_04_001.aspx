<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_001.aspx.vb" Inherits="WDAIIP.CP_04_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練機構資料</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
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
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" width="100%">
                        <tr>
                            <td>
                                <font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;</font><font class="font" color="#800000" size="2">訓練機構資料</font>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table1" cellspacing="1" cellpadding="1">
                        <tr>
                            <td style="width: 10%; height: 29px" class="bluecol_need">年度
                            </td>
                            <td style="height: 28px" class="whitecol">
                                <asp:DropDownList ID="yearlist" runat="server">
                                </asp:DropDownList>
                                <asp:RequiredFieldValidator ID="MustYear" runat="server" ErrorMessage="請選擇年度" Display="Dynamic" ControlToValidate="yearlist" CssClass="font"></asp:RequiredFieldValidator></FONT></FONT>
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">轄區
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="DistrictList" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">縣市
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="CityList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="8">
                                </asp:CheckBoxList>
                                <input id="CityHidden" type="hidden" value="0" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td style="width: 10%" class="bluecol">訓練計畫
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="PlanList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
                                </asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" runat="server">
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <table class="font" id="Table4" cellspacing="0" cellpadding="0" width="740" border="0">
            <tr>
                <td align="center">
                    <font face="新細明體">
                        <asp:Button ID="bt_search" runat="server" Width="60px" Text="明細查詢" CssClass="asp_button_M"></asp:Button>&nbsp;<asp:Button ID="bt_search1" runat="server" Width="60px" Text="統計查詢" CssClass="asp_button_M"></asp:Button>&nbsp;<asp:Button ID="bt_reset" runat="server" Width="60px" Text="重新設定" CssClass="asp_button_M"></asp:Button>&nbsp;<asp:Button ID="bt_export" runat="server" Width="60px" Text="資料匯出" CssClass="asp_button_M"></asp:Button></font>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
