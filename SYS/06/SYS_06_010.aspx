<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_010.aspx.vb" Inherits="WDAIIP.SYS_06_010" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>使用歷程查詢</title>
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
        <%--//決定date-picker元件使用的是西元年or民國年，by:20181019
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);--%>

        function clearDate(objId) {
            var myObj = document.getElementById(objId);
            if (myObj) { myObj.value = ""; }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;使用歷程查詢</asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td width="20%" class="bluecol">日期區間：</td>
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
                <td width="20%" class="bluecol">功能類別：</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlType1" runat="server" AutoPostBack="True">
                        <asp:ListItem Value="" Selected="True">全部</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">功能項目類別：</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlType2" runat="server">
                        <asp:ListItem Value="" Selected="True">全部</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">功能名稱：</td>
                <td class="whitecol">
                    <asp:TextBox ID="qFuncName" Width="50%" runat="server" MaxLength="50"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">操作類型：</td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlType3" runat="server">
                        <asp:ListItem Value="" Selected="True">全部</asp:ListItem>
                        <asp:ListItem Value="Insert">新增</asp:ListItem>
                        <asp:ListItem Value="Update">修改</asp:ListItem>
                        <asp:ListItem Value="Delete">刪除</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">關鍵字1：</td>
                <td class="whitecol">
                    <asp:TextBox ID="qKeyWord1" Width="50%" runat="server" MaxLength="120"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">表格：</td>
                <td class="whitecol">
                    <asp:TextBox ID="qtablename" Width="50%" runat="server" MaxLength="120"></asp:TextBox></td>
            </tr>

            <tr>
                <td width="20%" class="bluecol">帳號：</td>
                <td class="whitecol">
                    <asp:TextBox ID="qAcc" Width="30%" runat="server" MaxLength="30"></asp:TextBox></td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                    <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M" AuthType="QRY"></asp:Button>
                    <asp:Button ID="bt_export1" runat="server" Text="匯出" CssClass="asp_button_M" AuthType="EXP"></asp:Button>
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
                    <div runat="server" id="div1">
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn DataField="UserID" HeaderText="帳號" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="12%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="UserName" HeaderText="姓名" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="12%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="TRANSTYPE" HeaderText="操作類型" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="8%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="FUNCNAME" HeaderText="功能項目" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="20%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="TransTime" HeaderText="紀錄時間" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="12%"></asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="異動資料說明">
                                    <ItemStyle Width="36%" HorizontalAlign="LEFT"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:Label ID="hidCol1" runat="server" Text='<%# Bind("Conditions") %>' Visible="false"></asp:Label>
                                        <asp:Label ID="hidCol2" runat="server" Text='<%# Bind("BeforeValues") %>' Visible="false"></asp:Label>
                                        <asp:Label ID="hidCol3" runat="server" Text='<%# Bind("AfterValues") %>' Visible="false"></asp:Label>
                                        <asp:Label ID="lblInfo" runat="server" Text=""></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                    </div>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <input type="hidden" runat="server" id="hid_V" />
    </form>
</body>
</html>
