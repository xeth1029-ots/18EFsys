<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_013.aspx.vb" Inherits="WDAIIP.TR_04_013" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_04_013</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<font color="#990000">學員就業成果統計表_依縣市別</font>
                        </td>
                    </tr>
                </table>
                <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                    <tr>
                        <td class="whitecol">
                            結訓月數:
                            <asp:DropDownList ID="CPoint" runat="server" AutoPostBack="True">
                                <asp:ListItem Value="已結訓班級">已結訓班級</asp:ListItem>
                                <asp:ListItem Value="結訓滿三個月">結訓滿三個月</asp:ListItem>
                            </asp:DropDownList>
                            
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" ShowFooter="True">
                                <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                <Columns>
                                    <asp:BoundColumn DataField="CTName" HeaderText="縣市別" FooterText="合計"></asp:BoundColumn>
                                    <asp:BoundColumn DataField="InJobCount" HeaderText="就業人數" FooterText="0"></asp:BoundColumn>
                                    <asp:BoundColumn DataField="FinCount" HeaderText="結訓人數" FooterText="0"></asp:BoundColumn>
                                    <%--
									<asp:BoundColumn DataField="RejPeoCount" HeaderText="提前就業人數" FooterText="0"></asp:BoundColumn>
                                    --%>
                                    <asp:BoundColumn HeaderText="就業率"></asp:BoundColumn>
                                    <asp:BoundColumn DataField="TotalCost" HeaderText="預算金額" FooterText="0" DataFormatString="{0:#,##0.00}">
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                        <FooterStyle HorizontalAlign="Right"></FooterStyle>
                                    </asp:BoundColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            備註：
                            <br>
							<asp:Label ID="labMsg1" runat="server" ></asp:Label>
                            <br>
                            2.提前就業人數：學員實際參訓時數達總訓練時數1/2以上，經分署專案核定免負擔退訓賠償費用者。
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
