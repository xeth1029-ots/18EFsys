<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_001_01.aspx.vb" Inherits="WDAIIP.CP_04_001_01" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練機構資料</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" width="100%">
        <tr>
            <td>
                <font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;</font><font class="font" color="#800000" size="2">訓練機構資料</font>
            </td>
        </tr>
    </table>
    <table class="font" id="Table5" cellspacing="0" cellpadding="0" width="100%" border="0">
        <tr>
            <td>
                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" Width="100%">
                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                    <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                    <Columns>
                        <asp:BoundColumn HeaderText="序號"></asp:BoundColumn>
                        <asp:BoundColumn DataField="DistName" HeaderText="轄區"></asp:BoundColumn>
                        <asp:BoundColumn DataField="PlanName" HeaderText="計畫名稱"></asp:BoundColumn>
                        <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱"></asp:BoundColumn>
                        <asp:BoundColumn DataField="ComIDNO" HeaderText="統編"></asp:BoundColumn>
                        <asp:BoundColumn DataField="Address" HeaderText="地址"></asp:BoundColumn>
                        <asp:BoundColumn DataField="ContactName" HeaderText="聯絡人"></asp:BoundColumn>
                        <asp:BoundColumn DataField="Phone" HeaderText="電話"></asp:BoundColumn>
                        <asp:TemplateColumn Visible="False" HeaderText="功能">
                            <HeaderStyle Width="100px"></HeaderStyle>
                            <ItemTemplate>
                                <asp:Button runat="server" Text="詳細" CommandName="Company_Seqno" CausesValidation="false" ID="Button1"></asp:Button>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                    </Columns>
                    <PagerStyle Visible="False"></PagerStyle>
                </asp:DataGrid>
                <div align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </div>
            </td>
        </tr>
        <tr>
            <td style="height: 16px" align="center">
                <asp:Label ID="NoData" runat="server" CssClass="font"></asp:Label>
            </td>
        </tr>
        <tr>
            <td align="center">
                <asp:Button ID="Button2" runat="server" CausesValidation="False" Text="回上頁" CssClass="asp_button_S"></asp:Button>
            </td>
        </tr>
        <tr>
            <td align="left">
                <asp:Label ID="description" runat="server" CssClass="font">排序說明：以轄區、訓練計畫做排序</asp:Label>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
