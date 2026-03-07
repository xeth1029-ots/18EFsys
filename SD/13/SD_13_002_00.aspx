<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_13_002_00.aspx.vb" Inherits="WDAIIP.SD_13_002_00" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_13_002_00</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <%--<script language="javascript" src="../../js/date-picker2.js"></script>--%>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript">
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);

        function GETvalue() {
            document.getElementById('Button4').click();
        }

        function SetOneOCID() {
            document.getElementById('Button5').click();
        }

        function choose_class() {
            if (document.getElementById('OCID1').values == '') document.getElementById('Button5').click();
            document.getElementById('DataGridTable').style.display = 'none';
            openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value);
        }

        function CheckSearch() {
            if (document.getElementById('OCIDValue1').value == '') {
                alert('請選擇職類班別');
                return false;
            }
        }

        function SelectAll(flag, obj) {
            var MyTable = document.getElementById('DataGrid1');
            if (flag == 0) {
                //撥款全選
                var idx = obj.selectedIndex;
                for (i = 1; i < MyTable.rows.length; i++) {
                    MyTable.rows(i).cells(11).children(0).selectedIndex = idx;
                }
            } else {
                //撥款日期全選
                var objValue = MyTable.rows(1).cells(12).children(0).value;
                if (obj.checked && objValue != '') {
                    for (i = 1; i < MyTable.rows.length; i++) {
                        MyTable.rows(i).cells(12).children(0).value = objValue;
                    }
                } else if (!obj.checked) {
                    for (i = 2; i < MyTable.rows.length; i++) {
                        MyTable.rows(i).cells(12).children(0).value = '';
                    }
                }
            }
        }

        function selDate(idx) {
            var MyTable = document.getElementById('DataGrid1');
            show_calendar(MyTable.rows(idx).cells(12).children(0).id, '', '', 'CY/MM/DD');
        }

        function chkSave() {
            var msg = '';
            var MyTable = document.getElementById('DataGrid1');
            var AppliedStatus = 0;
            var AllotDate = '';
            for (i = 1; i < MyTable.rows.length; i++) {
                AllotDate = MyTable.rows(i).cells(12).children(0).value;
                if (AllotDate != '') {
                    if (!checkDate(AllotDate)) {
                        msg += '學號' + MyTable.rows(i).cells(0).innerHTML + '撥款日期格式有誤!\n';
                    }
                }
            }
            for (i = 1; i < MyTable.rows.length; i++) {
                AppliedStatus = MyTable.rows(i).cells(11).children(0).selectedIndex;
                AllotDate = MyTable.rows(i).cells(12).children(0).value;
                if (AppliedStatus == 1 && AllotDate == '') {
                    msg += '學號' + MyTable.rows(i).cells(0).innerHTML + '未填撥款日期!\n';
                }
                if (AppliedStatus == 0 && AllotDate != '') {
                    msg += '學號' + MyTable.rows(i).cells(0).innerHTML + '未選擇撥款狀況!\n';
                }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
    <style type="text/css">
        .auto-style1 { height: 30px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <asp:Button ID="Button5" Style="display: none" runat="server"></asp:Button><asp:Button ID="Button4" Style="display: none" runat="server"></asp:Button>
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--
                        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;補助金請領&gt;&gt;<FONT color="#990000">補助撥款</FONT></asp:Label>
                            </td>
                        </tr>
                    </table>
                    --%>
                    <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%">
                        <tbody>
                            <tr>
                                <td class="bluecol" width="100">訓練機構</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="center" runat="server" Width="410px"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
                                    <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">職類/班別</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                                    <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                    <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">經費審核</td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="AuditList" runat="server" RepeatDirection="Horizontal" Font-Size="X-Small"></asp:RadioButtonList></td>
                            </tr>
                            <tr>
                                <td class="whitecol" align="center" colspan="2">
                                    <asp:Button ID="btnSch" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button></td>
                            </tr>
                            <tr>
                                <td class="whitecol" align="center" colspan="2">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                            </tr>
                        </tbody>
                    </table>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowSorting="True">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="StudentID" SortExpression="StudentID" HeaderText="學號">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" SortExpression="IDNO" HeaderText="身分證號碼">
                                            <HeaderStyle ForeColor="#B0E2FF"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="是否取得&lt;BR&gt;結訓資格"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="出席達&lt;BR&gt;2/3"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="是否&lt;BR&gt;補助"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Total" HeaderText="總費用"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="SumOfMoney" HeaderText="補助費用"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="PayMoney" HeaderText="個人支付"></asp:BoundColumn>
                                        <asp:BoundColumn HeaderText="剩餘可&lt;BR&gt;用餘額"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="GovAppl2" HeaderText="其他申請&lt;BR&gt;中金額"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="撥款">
                                            <HeaderTemplate>
                                                撥款
                                            <asp:DropDownList ID="DropDownList1" runat="server">
                                                <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                                <asp:ListItem Value="已撥款">已撥款</asp:ListItem>
                                            </asp:DropDownList>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="AppliedStatus" runat="server">
                                                    <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                                    <asp:ListItem Value="已撥款">已撥款</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="撥款日期">
                                            <HeaderStyle Width="110px"></HeaderStyle>
                                            <HeaderTemplate>
                                                <asp:CheckBox ID="chkAll" runat="server"></asp:CheckBox>
                                                撥款日期
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtAllotDate" Width="80px" runat="server"></asp:TextBox>
                                                <asp:ImageButton ID="ibtDate" ImageUrl="../../images/show-calendar.gif" runat="server"></asp:ImageButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="撥款備註">
                                            <ItemTemplate>
                                                <asp:TextBox ID="AppliedNote" runat="server" TextMode="MultiLine"></asp:TextBox>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="auto-style1">
                                <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
