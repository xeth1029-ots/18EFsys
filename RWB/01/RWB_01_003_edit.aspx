<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="RWB_01_003_edit.aspx.vb" Inherits="WDAIIP.RWB_01_003_edit" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>資料下載</title>
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
    <form id="form1" method="post" runat="server" enctype="multipart/form-data">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;報名網維護&gt;&gt;資料下載</asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td width="20%" class="bluecol_need">類別：</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlType" runat="server" Width="30%" AutoPostBack="True">
                        <asp:ListItem Value="1" Selected="True">計畫表單</asp:ListItem>
                        <asp:ListItem Value="2">其他資料</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">計畫：</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlPlan" runat="server" Width="30%">
                        <asp:ListItem Value="" Selected="True">[無計畫內容]</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
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
                <td width="20%" class="bluecol_need">標題：</td>
                <td class="whitecol">
                    <asp:TextBox ID="txtTitle" Width="80%" runat="server" placeholder="請輸入連結標題"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">附件1上傳：</td>
                <td class="whitecol">
                    <asp:FileUpload ID="fu1" runat="server" Width="60%" />
                    <div id="divFile1" runat="server" visible="false">
                        <br />
                        &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="lnkF1Name" runat="server" Text="" ToolTip="F1"></asp:LinkButton>
                        &nbsp;<asp:Label ID="lblF1Name" runat="server" Visible="false"></asp:Label>
                        &nbsp;<asp:Label ID="lblF1Ext" runat="server" Visible="false"></asp:Label>
                    </div>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">附件2上傳：</td>
                <td class="whitecol">
                    <asp:FileUpload ID="fu2" runat="server" Width="60%" />
                    <div id="divFile2" runat="server" visible="false">
                        <br />
                        &nbsp;&nbsp;&nbsp;<asp:LinkButton ID="lnkF2Name" runat="server" Text="" ToolTip="F2"></asp:LinkButton>
                        &nbsp;<asp:Label ID="lblF2Name" runat="server" Visible="false"></asp:Label>
                        &nbsp;<asp:Label ID="lblF2Ext" runat="server" Visible="false"></asp:Label>
                    </div>
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
