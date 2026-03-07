<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_031.aspx.vb" Inherits="WDAIIP.SD_15_031" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>新住民參訓明細</title>
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;新住民參訓明細</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">年度</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddlYEARS_SCH" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區分署</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cblDISTID" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="cblDISTIDHidden" type="hidden" value="0" name="cblDISTIDHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓期間
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="STDATE_SCH1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDATE_SCH1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>～
                                <asp:TextBox ID="STDATE_SCH2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDATE_SCH2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓期間
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="FTDATE_SCH1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDATE_SCH1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>～
                                <asp:TextBox ID="FTDATE_SCH2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDATE_SCH2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
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
                                <asp:Label ID="Labmsg1" runat="server" Text=""></asp:Label>
                                <asp:Button ID="BTN_EXPORT1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
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
