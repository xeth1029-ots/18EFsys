<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_022.aspx.vb" Inherits="WDAIIP.TC_01_022" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>最近一次審查計分等級</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;最近一次審查計分等級</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="Table3" width="100%" cellpadding="1" cellspacing="1">
                        <tr>
                            <td colspan="2" align="center" class="table_title" width="100%">最近一次審查計分等級</td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">計畫年度</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="YEARS_ROC" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td width="20%" class="bluecol">申請階段</td>
                            <td class="whitecol" width="80%">
                                <asp:Label ID="APPSTAGE_N" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" align="center" class="table_title" width="100%">審查計分複審等級</td>
                        </tr>

                        <tr>
                            <td colspan="2" align="center">
                                <%--計畫年度、申請階段、分署、訓練單位、統一編號、審查計分等級--%>
                                <%--<asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="6%"></asp:BoundColumn>--%>
                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="DISTNAME" HeaderText="分署" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="15%" ItemStyle-Height="99pt">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練單位" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="15%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="COMIDNO" HeaderText="統一編號" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="11%">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="RLEVEL_2" HeaderText="審查計分等級" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="8%">
                                            <%--複審等級--%>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <%--<asp:TemplateColumn HeaderText="複審等級" HeaderStyle-Width="7%">
                                            <HeaderStyle Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                            <ItemTemplate>
                                                <asp:HiddenField ID="Hid_RLEVEL_2" runat="server" />
                                                <asp:DropDownList ID="ddlRLEVEL_2" runat="server"></asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="複審審核" HeaderStyle-Width="7%">
                                            <HeaderStyle Width="7%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                            <ItemTemplate>
                                                <asp:HiddenField ID="HidOSID2" runat="server" />
                                                <asp:HiddenField ID="Hid_SECONDCHKorg" runat="server" />
                                                <asp:DropDownList ID="ddlSECONDCHK" runat="server">
                                                    <asp:ListItem Selected="True" Value="">請選擇</asp:ListItem>
                                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>--%>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>

                        <tr>
                            <td class="whitecol" colspan="2" align="center">
                                <asp:Label ID="lab_msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>

        <asp:HiddenField ID="Hid_YEARS" runat="server" />
        <asp:HiddenField ID="Hid_APPSTAGE" runat="server" />
        <asp:HiddenField ID="Hid_ORGID" runat="server" />
        <asp:HiddenField ID="Hid_OTQID" runat="server" />
        <asp:HiddenField ID="HID_SCORINGID" runat="server" />

    </form>
</body>
</html>
