<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_015.aspx.vb" Inherits="WDAIIP.SD_05_015" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>結訓成績登錄</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button4').click();
        }
        function SetOneOCID() {
            document.getElementById('Button5').click();
        }
        function choose_class() {
            //if (document.getElementById('DataGridTable')) { document.getElementById('DataGridTable').style.display = 'none'; }

            var RIDValue = document.getElementById('RIDValue');
            var OCID1 = document.getElementById('OCID1');
            if (!OCID1) return;
            if (OCID1.value == '') { document.getElementById('Button5').click(); }
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }
        function CheckSearch() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (!OCIDValue1) return;
            if (OCIDValue1 && OCIDValue1.value == '') {
                alert('請選擇職類班別');
                return false;
            }
        }
        function ChangeAll(j) {
            var cells_num2 = 2;
            var MyTable = document.getElementById('DataGrid1');
            if (!MyTable) return;
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows[i].cells[cells_num2].children[0].selectedIndex = j;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;結訓成績登錄</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="55%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button5" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button4"></asp:Button>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="chooseclass" runat="server" onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowCustomPaging="true" CellPadding="3">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="StudentID" HeaderText="學號">
                                            <HeaderStyle Width="22%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle Width="22%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="33%" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderTemplate>
                                                是否取得結訓資格
                                            <asp:DropDownList ID="SelectAll" runat="server" CssClass="whitecol">
                                                <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                                <asp:ListItem Value="是">是</asp:ListItem>
                                                <asp:ListItem Value="否">否</asp:ListItem>
                                            </asp:DropDownList>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="CreditPoints" runat="server" CssClass="whitecol">
                                                    <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                                    <asp:ListItem Value="是">是</asp:ListItem>
                                                    <asp:ListItem Value="否">否</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False" Mode="NumericPages"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button3" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button6" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
