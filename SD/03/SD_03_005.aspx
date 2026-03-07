<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_005.aspx.vb" Inherits="WDAIIP.SD_03_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員刪除作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function SetOneOCID() {
            document.getElementById('Button4').click();
        }

        function search() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') {
                alert('請先選擇班級職類');
                return false;
            }
        }

        function choose_class() {
            //var RID = document.form1.RIDValue.value;
            var RIDValue = document.getElementById('RIDValue');
            var OCID1 = document.getElementById('OCID1');
            document.form1.TMID1.value = '';
            document.form1.TMIDValue1.value = '';
            document.form1.OCID1.value = '';
            document.form1.OCIDValue1.value = '';
            document.getElementById('msg').innerHTML = '';
            if (OCID1.value == '') { document.getElementById('Button4').click(); }
            openClass('../02/SD_02_ch.aspx?RWClass=1&RID=' + RIDValue.value);
        }

        function CheckData() {
            var MyTable = document.getElementById('DataGrid1');
            var msg = '';
            var Flag = false;
            for (i = 1; i < MyTable.rows.length; i++) {
                var checkboxObj1 = MyTable.rows[i].cells[0].children[0];
                var STUDID2Obj1 = MyTable.rows[i].cells[1]; //STUDID2
                var NameObj1 = MyTable.rows[i].cells[2]; //Name
                var snValue1 = '([' + STUDID2Obj1.innerHTML + ']' + NameObj1.innerHTML + ')\n';
                var DelResaonObj1 = MyTable.rows[i].cells[7].children[0]; //DelResaon
                var DelResaonObj2 = MyTable.rows[i].cells[7].children[1]; //DelReasonOther
                if (checkboxObj1.checked) {
                    Flag = true;
                    //DelResaon
                    if (DelResaonObj1.selectedIndex == 0) { msg += '請選擇原因' + snValue1; }
                    else if (DelResaonObj1.value == '3') {
                        //3:其他
                        if (DelResaonObj2.value == '') { msg += '請輸入其他原因' + snValue1; }
                        else {
                            if (DelResaonObj2.value.length > 50) { msg += '刪除原因不得輸入超過50個中文字元' + snValue1; }
                        }
                    }
                }
            }
            if (!Flag) msg += '請勾選學員\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            else {
                return confirm('您確定要刪除資料?')
            }
        }

        function SelectAll() {
            var MyTable = document.getElementById('DataGrid1');
            var checkedValue1 = MyTable.rows[0].cells[0].children[0].checked;
            for (i = 1; i < MyTable.rows.length; i++) {
                var checkboxObj1 = MyTable.rows[i].cells[0].children[0];
                if (!checkboxObj1.disabled) { checkboxObj1.checked = checkedValue1; }
            }
        }

        function SetChgDelRes1(obj1, DelReasonOther, Label1) {
            var DelReasonOther = document.getElementById(DelReasonOther);
            var Label1 = document.getElementById(Label1);
            var obj1 = document.getElementById(obj1);
            DelReasonOther.style.display = 'none';
            Label1.style.display = 'none';
            //alert(obj1.value);
            if (obj1.value == '3') {
                //3:其他
                DelReasonOther.style.display = '';
                Label1.style.display = '';
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
                                <%--<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;報到&gt;&gt;學員刪除作業</asp:Label>--%>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;學員刪除作業</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button8" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Button4" Style="display: none" runat="server"></asp:Button>
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班級 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button5" onclick="choose_class();" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol" colspan="2">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button></td>
                        </tr>
                    </table>
                    <%--<table width="100%"></table>--%>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <div style="overflow-y: auto; height: 800px">
                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5" />
                            <HeaderStyle CssClass="head_navy" />
                            <ItemStyle HorizontalAlign="Center" />
                            <Columns>
                                <asp:TemplateColumn HeaderText="選取">
                                    <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <HeaderTemplate><input type="checkbox" onclick="SelectAll();" /></HeaderTemplate>
                                    <ItemTemplate><input id="Checkbox1" type="checkbox" runat="server" /></ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="STUDID2" HeaderText="學號">
                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="NAME" HeaderText="姓名">
                                    <HeaderStyle Width="10%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="IDNO_MK" HeaderText="身分證號碼">
                                    <HeaderStyle Width="10%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Sex2" HeaderText="性別">
                                    <HeaderStyle Width="5%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="BIRTHDAY_MK" HeaderText="出生日期">
                                    <HeaderStyle Width="15%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="StudStatus2" HeaderText="學員狀態">
                                    <HeaderStyle Width="10%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="刪除原因">
                                    <HeaderStyle Width="35%" />
                                    <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                    <ItemTemplate>
                                        <asp:DropDownList ID="DelResaon" runat="server"></asp:DropDownList><br>
                                        <asp:TextBox ID="DelReasonOther" runat="server" Rows="3" TextMode="MultiLine"></asp:TextBox>
                                        <br />
                                        <asp:Label ID="Label1" runat="server" ForeColor="Red">最多只能輸入50個中文字</asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                        </asp:DataGrid>
                    </div>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button2" runat="server" Text="刪除學員" CssClass="asp_button_M"></asp:Button></td>
            </tr>
        </table>
    </form>
</body>
</html>
