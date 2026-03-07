 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="OB_01_010.aspx.vb" Inherits="WDAIIP.OB_01_010" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>OB_01_010</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table class="font" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <asp:Label ID="lblTitle1" runat="server"></asp:Label><asp:Label ID="lblTitle2" runat="server">
							<FONT face="新細明體">首頁&gt;&gt;委外訓練管理&gt;&gt;<font color="#990000">機構績效查詢</font></FONT>
                </asp:Label><font color="#000000">(<font face="新細明體"><font color="#ff0000">*</font>為必填欄位</font>)</font>
            </td>
        </tr>
    </table>
    <asp:Panel ID="panelSch" runat="server">
        <table class="font" border="0" cellspacing="1" cellpadding="1" width="740">
            <tr>
                <td>
                    <table class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td width="15%" class="bluecol">
                                機構名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="schOrgName" runat="server" Width="400px" MaxLength="20"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol">
                                廠商統編
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="schComIDNO" runat="server" Width="100px" MaxLength="15"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol">
                                統計年度
                            </td>
                            <td style="height: 19px" class="whitecol">
                                <asp:DropDownList ID="ddlyears" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <p align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button></p>
                </td>
            </tr>
            <tr>
                <td>
                    <table border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td align="center">
                                <p>
                                    <table id="DataGridTable" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td align="right">
                                                (為避免消耗主機效能，最大搜尋筆數為2000筆)共計：
                                                <asp:Label ID="RecordCount" runat="server"></asp:Label>&nbsp;筆資料
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" PagerStyle-Mode="NumericPages" PagerStyle-HorizontalAlign="Left" AllowSorting="True">
                                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="Years" HeaderText="年度">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="DistName" HeaderText="轄區">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="補登次數">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:LinkButton ID="aeCount" runat="server" CommandName="aeCount" ForeColor="Black" Text='<%# DataBinder.Eval(Container, "DataItem.aeCount") %>'>
                                                                </asp:LinkButton>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                    <PagerStyle Visible="False"></PagerStyle>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" PagerStyle-Mode="NumericPages" PagerStyle-HorizontalAlign="Left" AllowSorting="True">
                                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn HeaderText="序號">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="Years" HeaderText="年度">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="DistName" HeaderText="轄區">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="Account" HeaderText="授權帳號">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn HeaderText="開放功能">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="CreateDate" HeaderText="補登開始日">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="CreateDate" HeaderText="補登結束日">
                                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        </asp:BoundColumn>
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
                                </p>
                                <p>
                                    <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </asp:Panel>
    </form>
</body>
</html>
