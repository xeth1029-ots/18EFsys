<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_06_001_add.aspx.vb" Inherits="WDAIIP.CP_06_001_ADD" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CP_06_001_ADD</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script>
        function ChkData() {
            var msg = '';
            if (isEmpty('QuestNum')) { msg += '請輸入問卷代號\n'; }
            if (isEmpty('QuestName')) { msg += '請輸入問卷命名\n'; }

            if (msg == '') {
                return true;
            } else {
                alert(msg);
                return false;
            }
        }

        function ChkDataQ() {
            var msg = '';
            if (isEmpty('QuestNum')) { msg += '請輸入問卷代號\n'; }

            if (msg == '') {
                return true;
            } else {
                alert(msg);
                return false;
            }
        }				
    </script>
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
                                    首頁&gt;&gt;訓練查核與績效管理&gt;&gt;<font color="#990000">機構問卷設定-新增(修改)</font>
                                </td>
                            </tr>
                        </table>
                        <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                            <tr>
                                <td style="width: 100px" align="left" width="35" class="bluecol_need">
                                    問卷代號
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:TextBox ID="QuestNum" runat="server"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 100px" align="left" width="35" class="bluecol_need">
                                    問卷命名
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:TextBox ID="QuestName" runat="server"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 100px" align="left" width="35" class="bluecol_need">
                                    問卷歸屬
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:RadioButtonList ID="PathAttach" runat="server" Width="77%" CssClass="font" CellPadding="0" CellSpacing="0">
                                        <asp:ListItem Value="1" Selected="True">受評單位自評及評鑑委員評鑑表</asp:ListItem>
                                        <asp:ListItem Value="2">實地評鑑訪視紀錄表</asp:ListItem>
                                        <asp:ListItem Value="3">評鑑成果報告表</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td style="width: 100px" align="left" width="35" class="bluecol">
                                    是否啟用
                                </td>
                                <td bgcolor="#ffecec" class="whitecol">
                                    <asp:CheckBox ID="Useing" runat="server" Checked="True"></asp:CheckBox>
                                </td>
                            </tr>
                            <tr>
                                <asp:Panel ID="Panel1" runat="server" Width="100%">
                                    <td class="bluecol">
                                        排序方式<br>
                                        (查詢用)
                                    </td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="Order" runat="server" CellSpacing="0" CellPadding="0" CssClass="font" Width="77%" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="1" Selected="True">依序號排序</asp:ListItem>
                                            <asp:ListItem Value="2">依問題排序</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </asp:Panel>
                            </tr>
                        </table>
                        <p align="center">
                            <asp:Button ID="Query" runat="server" Width="134px" Text="[不儲存]查詢問卷題目" CssClass="asp_button_L"></asp:Button>&nbsp;
                            <asp:Button ID="Q_Set" runat="server" Width="145px" Text="[不儲存]問卷題目設定" CssClass="asp_button_L"></asp:Button>&nbsp;
                            <asp:Button ID="Save" runat="server" Text="儲存問卷" CssClass="asp_button_M"></asp:Button>&nbsp;
                            <asp:Button ID="return_btn" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                        </p>
                    </font>
                    <p align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></p>
                </td>
            </tr>
        </tbody>
    </table>
    <table id="DataGridTable_Main" style="width: 704px" cellspacing="1" cellpadding="1" width="704" border="0" runat="server">
        <tr>
            <td>
                <asp:DataGrid ID="DataGrid1" runat="server" Width="696px" CssClass="font" AutoGenerateColumns="False" AllowPaging="True">
                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                    <Columns>
                        <asp:BoundColumn DataField="Path" HeaderText="項次">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="Heading" HeaderText="題目文字">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="Seq" HeaderText="排序">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:TemplateColumn HeaderText="功能">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            <ItemTemplate>
                                <asp:Button ID="Edit" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                <asp:Button ID="Del" runat="server" Text="刪除" CommandName="del"></asp:Button>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:BoundColumn Visible="False" DataField="OGQID" HeaderText="OGQID"></asp:BoundColumn>
                        <asp:BoundColumn Visible="False" DataField="OGHID" HeaderText="OGHID"></asp:BoundColumn>
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
    </form>
</body>
</html>
