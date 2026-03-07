<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_13_001_95.aspx.vb" Inherits="TIMS.SD_13_001_95" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_13_001_95</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../style.css" type="text/css" rel="stylesheet">
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

        function ChangeMoney(num, obj, obj2, obj3, obj4) {
            var MyTable = document.getElementById('DataGrid1');
            if (isUnsignedInt(document.getElementById(obj).value)) {
                if (parseInt(document.getElementById(obj).value) > parseInt(document.getElementById(obj2).value)) {
                    document.getElementById(obj).select();
                    alert('超過剩餘補助金額,此學員的剩餘補助金額為' + document.getElementById(obj2).value);
                }
                else if (parseInt(document.getElementById(obj).value) > parseInt(document.getElementById(obj3).value)) {
                    document.getElementById(obj).select();
                    alert('超過最大補助金額,最大補助金額為' + document.getElementById(obj3).value);
                }
                else {
                    MyTable.rows(num).cells(8).innerHTML = MyTable.rows(num).cells(6).innerHTML - parseInt(document.getElementById(obj).value);
                    document.getElementById(obj4).value = MyTable.rows(num).cells(8).innerHTML;
                    var Total = parseInt(document.getElementById(obj2).value) - parseInt(document.getElementById(obj).value);
                    if (Total >= 0)
                        MyTable.rows(num).cells(9).innerHTML = Total;
                    else
                        MyTable.rows(num).cells(9).innerHTML = '<font color=Red>' + Total + '</font>';
                }
            }
            else {
                document.getElementById(obj).focus();
                alert('請輸入數字');
            }
        }

        function CheckData() {
            var MyTable = document.getElementById('DataGrid1');
            var msg = '';
            for (i = 1; i < MyTable.rows.length; i++) {
                var SumOfMoney = parseInt(MyTable.rows(i).cells(7).children(0).value);
                var RemainSub = parseInt(MyTable.rows(i).cells(7).children(1).value);
                var MaxSub = parseInt(MyTable.rows(i).cells(7).children(2).value);
                var MyCehck = MyTable.rows(i).cells(10).children(0).value;
                if (!MyCehck.disabled && MyCehck.checked) {
                    if (!isUnsignedInt(SumOfMoney)) {
                        msg += '請金額必須為數字(學員:' + MyTable.rows(i).cells(1).innerHTML + ')\n';
                    }
                    else {
                        if (SumOfMoney > RemainSub)
                            msg += '補助金額不能超過剩餘補助金額(學員:' + MyTable.rows(i).cells(1).innerHTML + ',剩餘金額' + RemainSub + ')\n';
                        else if (SumOfMoney > MaxSub)
                            msg += '補助金額不能超此班最大補助金額(學員:' + MyTable.rows(i).cells(1).innerHTML + ')\n';
                    }
                }
                if (msg != '') {
                    alert(msg);
                    return false;
                }
            }
        }
    </script>
</head>
<body ms_positioning="FlowLayout">
    <form id="form1" method="post" runat="server">
        <font face="新細明體">
            <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <%--
                            <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;補助金請領&gt;&gt;<FONT color="#990000">補助申請</FONT></asp:Label>
                                </td>
                            </tr>
                        </table>
                        --%>
                        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="600" border="0">
                            <tr>
                                <td class="SD_TD1" width="100"><font face="新細明體">&nbsp;&nbsp;&nbsp; 訓練機構</font></td>
                                <td class="SD_TD2">
                                    <asp:TextBox ID="center" runat="server" Width="310px"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" size="1" name="Hidden2" runat="server">
                                    <input id="Button2" type="button" value="..." name="Button2" runat="server"><br>
                                    <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button4"></asp:Button>
                                    <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()"><asp:Table ID="HistoryRID" runat="server" Width="310px"></asp:Table></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="SD_TD1"><font face="新細明體">&nbsp;&nbsp;&nbsp; 職類/班別<font color="red">*</font></font></td>
                                <td class="SD_TD2">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()"></asp:TextBox><asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()"></asp:TextBox><input onclick="choose_class()" type="button" value="..."><input id="TMIDValue1" type="hidden" size="1" name="TMIDValue1" runat="server"><input id="OCIDValue1" type="hidden" size="1" name="OCIDValue1" runat="server"><br>
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
                        <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>滑鼠移至姓名可以顯示此學員前五年的申請紀錄</td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font">
                                        <Columns>
                                            <asp:BoundColumn DataField="StudentID" HeaderText="學號"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="是否獲&lt;BR&gt;得學分">
                                                <ItemTemplate>
                                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="300px" CssClass="font" AutoGenerateColumns="False" BorderWidth="1px" BorderColor="Black" BackColor="LemonChiffon">
                                                        <Columns>
                                                            <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱"></asp:BoundColumn>
                                                            <asp:BoundColumn DataField="SumOfMoney" HeaderText="金額">
                                                                <HeaderStyle Width="40px"></HeaderStyle>
                                                            </asp:BoundColumn>
                                                            <asp:BoundColumn DataField="AppliedStatus" HeaderText="申請狀態">
                                                                <HeaderStyle Width="60px"></HeaderStyle>
                                                            </asp:BoundColumn>
                                                        </Columns>
                                                    </asp:DataGrid>
                                                    <asp:Label ID="CreditPoints" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn HeaderText="出席達&lt;BR&gt;2/3"></asp:BoundColumn>
                                            <asp:BoundColumn HeaderText="是否&lt;BR&gt;補助"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="Total" HeaderText="總費用"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="補助費用">
                                                <ItemTemplate>
                                                    <asp:TextBox ID="SumOfMoney" runat="server" Columns="8"></asp:TextBox><input id="RemainSub" type="hidden" size="1" runat="server"><input id="MaxSub" type="hidden" size="1" runat="server"><input id="PayMoney" type="hidden" size="1" runat="server">
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn HeaderText="個人支付"></asp:BoundColumn>
                                            <asp:BoundColumn HeaderText="剩餘可&lt;BR&gt;用餘額"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="是否提&lt;BR&gt;出申請">
                                                <ItemTemplate>
                                                    <input id="Checkbox1" type="checkbox" runat="server">
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn HeaderText="申請狀態"></asp:BoundColumn>
                                        </Columns>
                                    </asp:DataGrid></td>
                            </tr>
                            <tr>
                                <td align="center"><asp:Button ID="Button3" runat="server" Text="儲存"></asp:Button></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </font>
    </form>
</body>
</html>