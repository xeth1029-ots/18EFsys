<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_04_001_02.aspx.vb" Inherits="WDAIIP.CP_04_001_02" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_04_001_02</title>
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
                    <font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;</font><font
                        class="font" color="#800000" size="2">訓練機構資料</font>
                </td>
            </tr>
        </table>
        <table class="font" id="Table5" cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" AutoGenerateColumns="False"
                        Width="100%">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構名稱"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Class1" HeaderText="已開班">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class2" HeaderText="未開班">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class3" HeaderText="不開班">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Class4" HeaderText="總班數">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="student1" HeaderText="訓練總人數">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="student2" HeaderText="在訓總人數">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="student3" HeaderText="結訓總人數">
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                            </asp:BoundColumn>
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
                    <asp:Button ID="Button1" runat="server" Text="列印" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="Button3" runat="server" Text="匯出" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>&nbsp;<asp:Button
                        ID="Button2" runat="server" CausesValidation="False" Text="回上頁" CssClass="asp_button_S"></asp:Button>
                </td>
            </tr>
            <tr>
                <td align="left" style="height: 63px">
                    <p>
                        &nbsp;
                    </p>
                    <p>
                        <asp:Label ID="description" runat="server" CssClass="font" Width="384px">統計說明：依訓練機構不分計畫統計，必需有任何訓練資料才會列入統計 訓練總人數=>目前總參訓人數 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							在訓總人數=>目前學員狀態為在訓的總人數&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 結訓總人數=>目前學員狀態為結訓的總人數</asp:Label>
                    </p>
                </td>
            </tr>
        </table>
        </FONT>
    </form>
</body>
</html>
