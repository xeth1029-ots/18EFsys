<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_032.aspx.vb" Inherits="WDAIIP.SD_15_032" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>師資資料整合查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;師資資料整合查詢</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 18%">身分證號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="IDNO_SCH" runat="server" MaxLength="20" Width="60%"></asp:TextBox></td>
                            <td class="bluecol" style="width: 18%">姓名 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TEACHCNAME_SCH" runat="server" MaxLength="30" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">計畫年度</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddlYEARS_SCH1" runat="server">
                                </asp:DropDownList>至
                                <asp:DropDownList ID="ddlYEARS_SCH2" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">計畫</td>
                            <td class="whitecol" colspan="3">
                                <asp:CheckBoxList ID="cblTPLANID" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="HiddencblTPLANID" type="hidden" value="0" name="HiddencblTPLANID" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4" align="center">
                                <asp:Button ID="BTN_EXPORT1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4" align="center">
                                <asp:Label ID="Labmsg1" runat="server" CssClass="font"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <div align="center">
                        <asp:Label ID="msg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
