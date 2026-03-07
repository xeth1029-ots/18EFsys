<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_01_002_mq.aspx.vb" Inherits="WDAIIP.SYS_01_002_mq" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>帳號計畫賦予-查詢</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>

</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">

            <tr>
                <td>
                    <table id="AccountTable" runat="server" class="table_sch" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 18%">帳號</td>
                            <td class="whitecol" style="width: 32%"><asp:TextBox ID="txt_ACCTID" runat="server" MaxLength="50"></asp:TextBox></td>
                            <td class="bluecol" style="width: 18%">姓名</td>
                            <td class="whitecol" style="width: 32%"><asp:TextBox ID="txt_ACCNTNAME" runat="server" MaxLength="50"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">單位名稱</td>
                            <td class="whitecol"><asp:TextBox ID="txt_ORGNAME" runat="server" MaxLength="50" Columns="50"></asp:TextBox></td>
                            <td class="bluecol">查詢方式</td>
                            <td class="whitecol">全部使用LIKE：<asp:CheckBox ID="CB_LIKE11" runat="server" /></td>
                        </tr>
                        <tr>
                            <td class="bluecol">狀態</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="rdoIsUsed" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="Y" Selected="True">啟用中</asp:ListItem>
                                    <asp:ListItem Value="N">停用中</asp:ListItem>
                                </asp:RadioButtonList>
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
                                  <asp:Button ID="btn_SEARCH1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="btn_EXPORT1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4">
                                <asp:Label ID="labNAME1" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td colspan="4" class="whitecol">
                                <asp:DataGrid ID="DataGrid1" runat="server" HeaderStyle-CssClass="TD_TD1" Width="100%" 
                                    AutoGenerateColumns="False" Visible="False" CssClass="font" CellPadding="8">
                                    <AlternatingItemStyle BackColor="WhiteSmoke" />
                                    <HeaderStyle CssClass="head_navy" HorizontalAlign="Center"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號" ItemStyle-Width="6%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="YEARS" HeaderText="年度" ItemStyle-Width="10%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="DISTNAME" HeaderText="分署" ItemStyle-Width="16%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="PLANNAME" HeaderText="計畫名稱" ItemStyle-Width="16%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="單位名稱" ItemStyle-Width="16%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ISUSED_N" HeaderText="狀態" ItemStyle-Width="10%"></asp:BoundColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4" align="center">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4"  align="center">
                                <asp:Button ID="btnBACK1" runat="server" Text="回上頁" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
        </table>
    </form>
</body>
</html>
