<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_06_001.aspx.vb" Inherits="WDAIIP.CP_06_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>機構問卷設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
        <tbody>
            <tr>
                <td>
                    <font face="新細明體">
                        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    首頁&gt;&gt;訓練查核與績效管理&gt;&gt;<font color="#990000">機構問卷設定</font>
                                </td>
                            </tr>
                        </table>
                        <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                            <tr>
                                <td style="width: 100px; height: 42px" align="left" width="35" class="bluecol">
                                    機構
                                </td>
                                <td style="height: 42px" bgcolor="#ffecec" class="whitecol">
                                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()"></asp:TextBox><input id="Button5"
                                        type="button" value="..." name="Button5" runat="server" 
                                        class="button_b_Mini"><input id="RIDValue" type="hidden"
                                            name="Hidden1" runat="server"><br>
                                    <span id="HistoryList2" style="display: none; position: absolute">
                                        <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                        </asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 100px" align="left" width="35" class="bluecol">
                                    問卷編號
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:TextBox ID="QuestNum" runat="server"></asp:TextBox>
                                </td>
                            </tr>
                        </table>
                        <p align="center">
                            <asp:Button ID="Query" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
                            <asp:Button ID="Btn_Add" runat="server" Text="新增" ToolTip="新增前請先選好機構" 
                                CssClass="asp_button_S"></asp:Button>
                        </p>
                    </font>
                    <p align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></p>
                </td>
            </tr>
        </tbody>
    </table>
    <table id="Table4" style="width: 704px" cellspacing="1" cellpadding="1" width="704"
        border="0" runat="server">
        <tr>
            <td>
                <asp:DataGrid ID="DataGrid_Main" runat="server" Width="696px" CssClass="font" AutoGenerateColumns="False"
                    AllowPaging="True">
                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                    <Columns>
                        <asp:BoundColumn DataField="QuestNum" HeaderText="問卷編號">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="QuestName" HeaderText="問卷名稱">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="Useing" HeaderText="啟用狀態">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:TemplateColumn HeaderText="功能">
                            <HeaderStyle HorizontalAlign="Center" Width="120px"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            <ItemTemplate>
                                <asp:Button ID="Edit" runat="server" Width="34px" Text="修改" CommandName="edit"></asp:Button>
                                <asp:Button ID="Del" runat="server" Width="34px" Text="刪除" CommandName="del"></asp:Button>
                                <asp:Button ID="View" runat="server" Width="34px" Text="檢視" CommandName="view"></asp:Button>
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
    </FONT>
    <p align="center">
        &nbsp;</p>
    </TD></TR></TBODY></TABLE></form>
</body>
</html>
