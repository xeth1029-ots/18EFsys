<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="RWB_01_002_edit.aspx.vb" Inherits="WDAIIP.RWB_01_002_edit" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>相關連結</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <%--<script type="text/javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <%--<script type="text/javascript" language="javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181019
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);
    </script>--%>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;報名網維護&gt;&gt;相關連結</asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td width="20%" class="bluecol_need">上稿日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtCDATE1" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">上架日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtSDATE1" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span1" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtSDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span2" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= txtSDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    <asp:DropDownList ID="ddlC_SDATE_hh1" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlC_SDATE_mm1" runat="server"></asp:DropDownList>分
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">連結類型：</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblType" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="1" Selected="True">訓練資源連結</asp:ListItem>
                        <asp:ListItem Value="2">求職求才資源連結</asp:ListItem>
                        <asp:ListItem Value="3">其他資源連結</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">連結標題：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtTitle" Width="80%" runat="server" placeholder="請輸入連結標題"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">連結網址：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtURL" Width="80%" runat="server" placeholder="請輸入連結網址"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">停用日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtEDATE1" Width="20%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span3" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtEDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span4" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= txtEDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    <asp:DropDownList ID="ddlC_EDATE_hh1" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlC_EDATE_mm1" runat="server"></asp:DropDownList>分
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_S" AuthType="SAVE"></asp:Button>
                    <asp:Button ID="bt_cancle" runat="server" Text="取消" CausesValidation="False" CssClass="asp_button_S" AuthType="CANCLE"></asp:Button>&nbsp;
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
        </table>
        <input type="hidden" runat="server" id="hid_V" />
    </form>
</body>
</html>
