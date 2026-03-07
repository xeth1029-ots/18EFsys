<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_003_other.aspx.vb" Inherits="WDAIIP.SD_02_003_other" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>錄訓作業-挑選其他志願</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function chall() {
            var docform1 = document.form1;
            docform1.SETID.checked = docform1.choose.checked
            for (var i = 0; i < docform1.SETID.length; i++)
                docform1.SETID[i].checked = docform1.choose.checked
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" class="table_nw">
            <tr>
                <td class="whitecol">
                    <div style="margin-top: 3px; margin-bottom: 3px">
                        <asp:Label ID="classname" runat="server"></asp:Label>
                    </div>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="false" CssClass="font">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn>
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <HeaderTemplate>
                                    <input type="checkbox" name="choose" onclick="chall();" />
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <input type="checkbox" name="setid" value='<%#databinder.eval(container.dataitem,"setid")%>@#<%#databinder.eval(container.dataitem,"enterdate","{0:d}")%>@#<%#databinder.eval(container.dataitem,"sernum")%>' />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="examno" HeaderText="准考證序號">
                                <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="name" HeaderText="姓名">
                                <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="writeresult" HeaderText="筆試成績">
                                <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="oralresult" HeaderText="口試成績">
                                <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="totalresult" HeaderText="總成績">
                                <HeaderStyle HorizontalAlign="Center" Width="13%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="SELRESULTID1" HeaderText="志願一">
                                <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="WISH_TXT" HeaderText="本班志願">
                                <HeaderStyle HorizontalAlign="Center" Width="14%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn Visible="false" DataField="setid"></asp:BoundColumn>
                            <asp:BoundColumn Visible="false" DataField="enterdate"></asp:BoundColumn>
                            <asp:BoundColumn Visible="false" DataField="sernum"></asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                    <div align="center" class="whitecol">
                        <asp:Button ID="button1" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center" class="whitecol">
                        <asp:Label ID="msg" runat="server" ForeColor="red" CssClass="font"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
