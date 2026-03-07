<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_13_History.aspx.vb" Inherits="WDAIIP.SD_13_History" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>補助審核-查詢重複參訓</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <style type="text/css">
        body { margin: 20px; overflow-y: auto; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table>
            <tr>
                <td align="center">
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn HeaderText="序號"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                            <%--<asp:BoundColumn DataField="DISTNAME" HeaderText="轄區中心"></asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區分署"></asp:BoundColumn>
                            <asp:BoundColumn DataField="planname" HeaderText="訓練計畫"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
                            <%--<asp:BoundColumn DataField="Classname" HeaderText="班級名稱"></asp:BoundColumn>--%>
                            <asp:TemplateColumn HeaderText="班級名稱">
                                <ItemTemplate>
                                    <asp:Label ID="lab_Classname" runat="server"></asp:Label>
                                    <asp:LinkButton ID="lib_Classname" runat="server" CommandName="Link1"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="Sfdate" HeaderText="受訓期間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="WEEKS" HeaderText="上課時間">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            </asp:BoundColumn>
                            <%--<asp:BoundColumn DataField="DISTNAME2" HeaderText="重複參訓-轄區中心"></asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="DISTNAME2" HeaderText="重複參訓-轄區分署"></asp:BoundColumn>
                            <asp:BoundColumn DataField="planname2" HeaderText="重複參訓-訓練計畫"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName2" HeaderText="重複參訓-訓練機構"></asp:BoundColumn>
                            <%--<asp:BoundColumn DataField="Classname2" HeaderText="重複參訓-班級名稱"></asp:BoundColumn>--%>
                            <asp:TemplateColumn HeaderText="重複參訓-班級名稱">
                                <ItemTemplate>
                                    <asp:Label ID="lab_Classname2" runat="server"></asp:Label>
                                    <asp:LinkButton ID="lib_Classname2" runat="server" CommandName="Link2"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="Sfdate2" HeaderText="重複參訓-受訓期間"></asp:BoundColumn>
                            <asp:BoundColumn DataField="WEEKS2" HeaderText="上課時間">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="STUDSTATUS_N1" HeaderText="訓練狀態">
                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            </asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_ROCYEARS" runat="server" />
        <asp:HiddenField ID="Hid_OCID1" runat="server" />
        <%--<asp:HiddenField ID="Hid_OCID2" runat="server" />--%>
    </form>
</body>
</html>
