<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_011.aspx.vb" Inherits="WDAIIP.SYS_06_011" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>登入紀錄查詢</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <%--<script type="text/javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181019
      <%--  var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);--%>

        function clearDate(objId) {
            var myObj = document.getElementById(objId);
            myObj.value = "";
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;登入紀錄查詢</asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td width="20%" class="bluecol">日期區間</td>
                <td class="whitecol">
                    <asp:TextBox ID="qDATE1" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span1" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= qDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span2" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= qDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="qDATE2" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span id="span3" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= qDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span4" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= qDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">類型</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlType" runat="server">
                        <asp:ListItem Value="" Selected="True">全部</asp:ListItem>
                        <asp:ListItem Value="LOGIN">登入</asp:ListItem>
                        <asp:ListItem Value="LOGOUT">登出</asp:ListItem>
                        <asp:ListItem Value="LOGINE1">登入失敗</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">帳號</td>
                <td class="whitecol">
                    <asp:TextBox ID="qAcc" Width="40%" runat="server" MaxLength="50"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">姓名</td>
                <td class="whitecol">
                    <asp:TextBox ID="qUserName" Width="40%" runat="server" MaxLength="50"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">紀錄時間
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlAccTime1_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlAccTime1_MM" runat="server"></asp:DropDownList>分
                    ～
                    <asp:DropDownList ID="ddlAccTime2_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlAccTime2_MM" runat="server"></asp:DropDownList>分
                </td>
            </tr>
            <tr>
                <td class="bluecol" style="width: 20%">匯出年月
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="DDL_YEAR1" runat="server"></asp:DropDownList>年
                    <asp:DropDownList ID="DDL_MONTH1" runat="server"></asp:DropDownList>月
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">20</asp:TextBox>
                    <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M" AuthType="QRY"></asp:Button>
                    <asp:Button ID="BTN_EXP1" runat="server" Text="匯出每月登入異常日誌" CssClass="asp_Export_M" />
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
        </table>
        <table id="tb_Sch" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="UserID" HeaderText="帳號">
                                <ItemStyle HorizontalAlign="Center" Width="14%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="UserName" HeaderText="姓名">
                                <ItemStyle HorizontalAlign="Center" Width="16%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TYPE" HeaderText="類型">
                                <ItemStyle HorizontalAlign="Center" Width="8%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn>
                                <HeaderTemplate>紀錄時間</HeaderTemplate>
                                <ItemStyle HorizontalAlign="Center" Width="12%"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="Lab_AccessTime" runat="server" Text=""></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <%--<asp:BoundColumn DataField="AccessTime" HeaderText="紀錄時間">
                                <ItemStyle HorizontalAlign="Center" Width="12%"></ItemStyle>
                            </asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="RemoteIP" HeaderText="IP">
                                <ItemStyle HorizontalAlign="Center" Width="16%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ResultMessage" HeaderText="訊息">
                                <ItemStyle HorizontalAlign="Center" Width="22%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="myBrowserInfo" HeaderText="使用瀏覽器">
                                <ItemStyle HorizontalAlign="Center" Width="12%"></ItemStyle>
                            </asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <input type="hidden" runat="server" id="hid_V" />
        <asp:HiddenField ID="Hid_DTV1" runat="server" />
    </form>
</body>
</html>
