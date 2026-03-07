<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_13_002_95.aspx.vb" Inherits="TIMS.SD_13_002_95" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_13_002_95</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script src="../../js/common.js"></script>
    <script>
        function GETvalue() {
            document.getElementById('Button4').click();
        }

        function choose_class() {
            document.getElementById('DataGridTable').style.display = 'none';
            openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value);
        }

        function CheckSearch() {
            if (document.getElementById('OCIDValue1').value == '') {
                alert('請選擇職類班別');
                return false;
            }
        }

        function SelectAll(idx) {
            var MyTable = document.getElementById('DataGrid1')
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows(i).cells(10).children(0).selectedIndex = idx;
            }
        }
    </script>
</head>
<body ms_positioning="FlowLayout">
    <form id="form1" method="post" runat="server">
        <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button4"></asp:Button>
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%"
            border="0">
            <tr>
                <td>
                    <%--
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;補助金請領&gt;&gt;<FONT color="#990000">補助查核</FONT></asp:Label>
                            </td>
                        </tr>
                    </table>
                    --%>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="SD_TD1" width="100"><font face="新細明體">&nbsp;&nbsp;&nbsp; 訓練機構</font></td>
                            <td class="SD_TD2">
                                <asp:TextBox ID="center" runat="server" Width="310px"></asp:TextBox>
                                <input id="RIDValue" type="hidden" size="1" name="RIDValue" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server"><br>
                                <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()"><asp:Table ID="HistoryRID" runat="server" Width="310px"></asp:Table></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="SD_TD1">&nbsp;&nbsp;&nbsp; 職類/班別<font color="red">*</font></td>
                            <td class="SD_TD2">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="...">
                                <input id="TMIDValue1" type="hidden" size="1" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" size="1" name="OCIDValue1" runat="server"><br>
                                <span id="HistoryList" style="display: none; left: 270px; position: absolute"><asp:Table ID="HistoryTable" runat="server" Width="310"></asp:Table></span>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2"><asp:Button ID="Button1" runat="server" Text="查詢"></asp:Button></td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2"><asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                    </table>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                    <Columns>
                                        <asp:BoundColumn DataField="StudentID" HeaderText="學號"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="是否獲&lt;BR&gt;得學分"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="出席達&lt;BR&gt;2/3"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="是否&lt;BR&gt;補助"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Total" HeaderText="總費用"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="SumOfMoney" HeaderText="補助費用"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="PayMoney" HeaderText="個人支付"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="剩餘可&lt;BR&gt;用餘額"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="審核">
                                            <HeaderTemplate>
                                                審核
                                            <asp:DropDownList ID="DropDownList1" runat="server">
                                                <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                                <asp:ListItem Value="通過">通過</asp:ListItem>
                                                <asp:ListItem Value="不通過">不通過</asp:ListItem>
                                            </asp:DropDownList>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="AppliedStatus" runat="server">
                                                    <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                                    <asp:ListItem Value="通過">通過</asp:ListItem>
                                                    <asp:ListItem Value="不通過">不通過</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="審核備註">
                                            <ItemTemplate>
                                                <asp:TextBox ID="AppliedNote" runat="server" TextMode="MultiLine"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center"><asp:Button ID="Button3" runat="server" Text="儲存"></asp:Button></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>