<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_014_A.aspx.vb" Inherits="WDAIIP.CM_03_014_A" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>原住民結訓人數統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
		 
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
						首頁&gt;&gt;訓練與需求管理&gt;&gt;統計分析&gt;&gt;<FONT color="#990000">原住民結訓人數統計表</FONT>
                    </asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" ShowFooter="True" CssClass="font" AutoGenerateColumns="False" Width="100%" AllowCustomPaging="True" AllowPaging="True" PageSize="99">
                        <FooterStyle HorizontalAlign="Center" CssClass="bluecol_sub"></FooterStyle>
                        <AlternatingItemStyle HorizontalAlign="Center" BackColor="#f5f5f5"></AlternatingItemStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:TemplateColumn HeaderText="計畫名稱" FooterText="合計">
                                <ItemTemplate>
                                    <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue" CommandName="searchB">LinkButton</asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="trainClass" HeaderText="參訓班數"></asp:BoundColumn>
                            <asp:BoundColumn DataField="trainNum" HeaderText="訓練人數"></asp:BoundColumn>
                            <asp:BoundColumn DataField="closeNum" HeaderText="結訓人數"></asp:BoundColumn>
                            <asp:BoundColumn DataField="JobNum" HeaderText="就業人數"></asp:BoundColumn>
                            <asp:BoundColumn DataField="NJobNum" HeaderText="不就業人數"></asp:BoundColumn>
                            <asp:BoundColumn DataField="xJobNum" HeaderText="未就業人數"></asp:BoundColumn>
                        </Columns>
                        <PagerStyle HorizontalAlign="Center" Position="Top"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td style="height: 25px" align="center" colspan="2">
                    <font face="新細明體">
                        <asp:Button ID="Button3" runat="server" Text="回上層" CssClass="asp_Export_M"></asp:Button></font>
                </td>
            </tr>
        </table>
        <input id="hYears" type="hidden" name="hYears" runat="server">
        <input id="hSTDate1" type="hidden" name="hSTDate1" runat="server">
        <input id="hSTDate2" type="hidden" name="hSTDate2" runat="server">
        <input id="hFTDate1" type="hidden" name="hFTDate1" runat="server">
        <input id="hFTDate2" type="hidden" name="hFTDate2" runat="server">
        <input id="hDistID1" type="hidden" name="hDistID1" runat="server">
        <input id="hTPlanID1" type="hidden" name="hTPlanID1" runat="server">
        <input id="hBudgetID" type="hidden" name="hBudgetID" runat="server">
    </form>
</body>
</html>
