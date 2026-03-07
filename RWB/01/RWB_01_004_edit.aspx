<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="RWB_01_004_edit.aspx.vb" Inherits="WDAIIP.RWB_01_004_edit" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>Q＆A</title>
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;報名網維護&gt;&gt;Q＆A</asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
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
                <td width="20%" class="bluecol_need">問題類型：</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblType" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="1" Selected="True">產業人才投資方案</asp:ListItem>
                        <asp:ListItem Value="2">自辦在職訓練</asp:ListItem>
                        <asp:ListItem Value="3">企業委託訓練</asp:ListItem>
                        <asp:ListItem Value="4">充電起飛</asp:ListItem>
                        <asp:ListItem Value="5">網站操作問題</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">問題內容：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtQ" Width="80%" runat="server" placeholder="請輸入問題內容" TextMode="MultiLine" Rows="8"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">回答內容：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtA" Width="80%" runat="server" placeholder="請輸入回答內容" TextMode="MultiLine" Rows="14"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">排序序號：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtCSORT1" Width="20%" runat="server" placeholder="請輸入序號"></asp:TextBox>
                    <asp:Label ID="lab_msg_CSORT1" runat="server" ForeColor="Red" Text="(由小到大排序)"></asp:Label>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">啟用狀態：</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblUse" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="Y" Selected="True">啟用</asp:ListItem>
                        <asp:ListItem Value="N">停用</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
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
