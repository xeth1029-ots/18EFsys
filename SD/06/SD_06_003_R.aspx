<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_06_003_R.aspx.vb" Inherits="WDAIIP.SD_06_003_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_06_003_R</title>
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
            document.getElementById('btnGetOneClass').click();
        }
        function CheckSearch() {
            if (document.getElementById('OCIDValue1').value == '') {
                alert('請選擇班級');
                return false;
            }
        }
        function SelectAll(num, Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows(i).cells(num).children(0).disabled == false) {
                    MyTable.rows(i).cells(num).children(0).checked = Flag;
                }
            }
        }
        function CheckPrint(num) {
            var MyTable = document.getElementById('DataGrid1');
            var SOCID = '';
            //var SOCID2='';
            //var YN='';
            for (i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows(i).cells(num).children(0).checked) {
                    if (SOCID != '') { SOCID += ','; }
                    SOCID += MyTable.rows(i).cells(num).children(0).value;
                }

            }

            if (SOCID == '') {
                alert('請選擇學員');
                return false;
            }
            else {
                if (num == 0) {
                    openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=SD_06_003_R&SOCID=' + SOCID
						+ '&Years=' + document.getElementById('Years').value + '&PMonth=' + document.getElementById('PMonth').value
						+ '&PDay=' + document.getElementById('PDay').value + '&PNO=' + document.getElementById('PNO').value);
                }
                else {
                    openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=SD_06_003_R_2&SOCID=' + SOCID
						+ '&Years=' + document.getElementById('Years').value + '&PMonth=' + document.getElementById('PMonth').value
						+ '&PDay=' + document.getElementById('PDay').value + '&PNO=' + document.getElementById('PNO').value);
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;學員動態管理&gt;&gt;加退保管理&gt;&gt;<font color="#990000">加退保申請表列印</font>
                            </td>
                        </tr>
                    </table>
                    <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="100">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" type="button" value="..." name="Button8" runat="server" class="asp_button_Mini" /><br />
                                <asp:Button ID="btnGetOneClass" Style="display: none" runat="server" Text="btnGetOneClass"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                                <input id="Button5" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" /><br />
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td colspan="2">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="50px"></HeaderStyle>
                                            <HeaderTemplate>
                                                <input onclick="SelectAll(0, this.checked);" type="checkbox">加保
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="Checkbox1" type="checkbox" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="50px"></HeaderStyle>
                                            <HeaderTemplate>
                                                <input type="checkbox" onclick="SelectAll(1, this.checked);">退保
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="Checkbox2" type="checkbox" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="StudentID" HeaderText="學號"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Birthday" HeaderText="出生日期" DataFormatString="{0:d}"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ApplyInsurance" HeaderText="加保日期" DataFormatString="{0:d}"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="DropoutInsurance" HeaderText="退保日期" DataFormatString="{0:d}"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="InsureSalary" HeaderText="投保薪資"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="曾否領取老&lt;br&gt;年或三等以&lt;br&gt;上殘廢給付">
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:DropDownList ID="OtherSubsidy" runat="server">
                                                    <asp:ListItem Value="N">否</asp:ListItem>
                                                    <asp:ListItem Value="Y">有</asp:ListItem>
                                                </asp:DropDownList>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                            <td></td>
                        </tr>
                        <tr>
                            <td align="right">報表日期:&nbsp;
                            </td>
                            <td align="left">民國
                            <asp:TextBox ID="Years" runat="server" Width="40px"></asp:TextBox>年
                            <asp:DropDownList ID="PMonth" runat="server">
                            </asp:DropDownList>
                                月
                            <asp:DropDownList ID="PDay" runat="server">
                            </asp:DropDownList>
                                日
                            </td>
                        </tr>
                        <tr>
                            <td align="right" width="40%">報表編號:&nbsp;
                            </td>
                            <td align="left" width="60%">
                                <asp:TextBox ID="PNO" runat="server" Width="224px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2">
                                <asp:Button ID="Save" runat="server" Text="存檔" CssClass="asp_button_S"></asp:Button>
                                <input id="PrintA" type="button" value="列印加保申報表" runat="server" class="asp_Export_M" />
                                <input id="PrintB" type="button" value="列印退保申報表" runat="server" class="asp_Export_M" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
