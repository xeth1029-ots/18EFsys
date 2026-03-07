<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_001.aspx.vb" Inherits="WDAIIP.SYS_01_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>帳號設定</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
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
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;帳號設定</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">帳號 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="nameid" MaxLength="22" runat="server" Width="60%"></asp:TextBox></td>
                            <td class="bluecol" style="width: 20%">姓名 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="namefield" MaxLength="15" runat="server" Width="40%"></asp:TextBox></td>
                        </tr>
                        <%--<tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="50%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>--%>
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練單位 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TBplan" MaxLength="100" runat="server" Width="50%"></asp:TextBox>
                                &nbsp;<input id="choice_button" type="button" value="選擇" name="choice_button" runat="server" class="asp_button_M">
                                &nbsp;<asp:Button ID="btnClear1" runat="server" Text="清除" CssClass="asp_button_M"></asp:Button>
                                <span id="HistoryList2" style="position: relative; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">啟用 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="isused" runat="server">
                                    <asp:ListItem Value="">==請選擇==</asp:ListItem>
                                    <asp:ListItem Value="Y">啟用</asp:ListItem>
                                    <asp:ListItem Value="N">不啟用</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">角色 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="Role" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">最後登入時間
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="LastDATE1" runat="server" Columns="8" Width="15%" MaxLength="10"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('LastDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                <asp:DropDownList ID="ddlLastDATE1_HH" runat="server"></asp:DropDownList>時
                                <asp:DropDownList ID="ddlLastDATE1_MM" runat="server"></asp:DropDownList>分
                                ～
                                <asp:TextBox ID="LastDATE2" runat="server" Columns="8" Width="15%" MaxLength="10"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('LastDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                <asp:DropDownList ID="ddlLastDATE2_HH" runat="server"></asp:DropDownList>時
                                <asp:DropDownList ID="ddlLastDATE2_MM" runat="server"></asp:DropDownList>分
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td colspan="4" class="whitecol" align="center" width="100%"><%--asp_Export_M / asp_button_M--%>
                                <asp:Button ID="but_search" runat="server" Text="查詢" CssClass="asp_button_M" AuthType="QRY"></asp:Button>
                                &nbsp;<asp:Button ID="but_add" runat="server" Text="新增" CssClass="asp_button_M" AuthType="ADD"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AllowSorting="True" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="Account" SortExpression="account" HeaderText="帳號">
                                            <ItemStyle HorizontalAlign="Center" Width="16%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="name" HeaderText="姓名">
                                            <ItemStyle HorizontalAlign="Center" Width="16%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="RoleName" SortExpression="RoleID" HeaderText="角色">
                                            <ItemStyle HorizontalAlign="Center" Width="16%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="所屬單位">
                                            <ItemStyle HorizontalAlign="Center" Width="22%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="isused" HeaderText="啟用狀態">
                                            <ItemStyle HorizontalAlign="Center" Width="8%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="LastDATE" HeaderText="最後登入時間">
                                            <ItemStyle HorizontalAlign="Center" Width="15%" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <ItemStyle HorizontalAlign="Center" Width="10%" />
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnEdit" runat="server" Text="修改" CssClass="linkbutton" CommandName="btnEdit" AuthType="UPD"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                    <center>
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></center>
                </td>
            </tr>
        </table>
        <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
        <input id="hidPlanID" type="hidden" name="hidPlanID" runat="server" />
        <%--<input id="Orgidvalue" type="hidden" name="Orgidvalue" runat="server" />--%>
        <%--<input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server" />--%>
    </form>
</body>
</html>
