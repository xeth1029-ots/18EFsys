<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_024.aspx.vb" Inherits="WDAIIP.SYS_03_024" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>單位群組設定</title>
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
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">
				    首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;單位群組設定
                    </asp:Label>
                </td>
            </tr>
        </table>

        <%--<table class="font" width="740">
        <tr>
            <td class="font">
                首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;<font color="#990000"><font color="#990000">單位群組設定<font face="Times New Roman" size="2"></font></font></font>
            </td>
        </tr>
    </table>--%>
        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
            <tr>
                <td class="bluecol" style="width: 20%">年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="list_Years" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">轄區
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="list_DistID" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">計畫代碼
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="list_PlanID" runat="server" AutoPostBack="true">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練單位
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="list_OrgID" AutoPostBack="true" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">功能選項
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblType" runat="server" RepeatDirection="Horizontal" CssClass="font">
                        <asp:ListItem Value="0">新增群組</asp:ListItem>
                        <asp:ListItem Value="1">刪除群組</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
        </table>
        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server" id="table2">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid3" CssClass="font" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="選用" ItemStyle-Width="6%" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:CheckBox ID="chk_GroupValid" runat="server"></asp:CheckBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="建檔單位" ItemStyle-Width="22%" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:Label ID="lab_DistName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="群組階層" ItemStyle-Width="12%" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:Label ID="lab_TypeName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="群組名稱">
                                <HeaderStyle Width="35%"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lab_GroupName" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="備註">
                                <HeaderStyle Width="25%"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:Label ID="lab_GroupNote" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle HorizontalAlign="Right" Mode="NumericPages"></PagerStyle>
                    </asp:DataGrid><asp:Label ID="lab_Msg2" runat="server" ForeColor="Red" Visible="False">無群組可賦予</asp:Label>
                </td>
            </tr>
            <tr id="tr_btn" runat="server">
                <td align="center" class="whitecol">
                    <asp:Button ID="btn_SaveGroup" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="btn_CancelGroup" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
