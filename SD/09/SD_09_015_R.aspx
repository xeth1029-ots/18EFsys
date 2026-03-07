<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_015_R.aspx.vb" Inherits="WDAIIP.SD_09_015_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_09_015_R</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <%--link href="../../style.css" type="text/css" rel="stylesheet"--%>
    <link href="../../css/style.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button12').click();
        }
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?RID=' + RID);
        }
        function ReportPrint() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇班級職類\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
            return true;
        }
        function ChangeAll(obj) {
            var objLen = document.form1.length;
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (obj.checked == true) {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        if (document.form1.elements[iCount].disabled == false) {
                            document.form1.elements[iCount].checked = true;
                        }
                    }
                }
                else {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        document.form1.elements[iCount].checked = false;
                    }
                }
            }
        }

        function CheckPrint(type) {
            var flag = false;
            var MyTable = document.getElementById('DataGrid1');
            var OCID = '';
            //var OCID=document.getElementById('OCIDValue1').value;

            for (var i = 1; i < MyTable.rows.length; i++) {
                var MyCheck = false;
                var Mycells0 = MyTable.rows[i].cells[0];
                if (Mycells0.children[2] == null) {
                    MyCheck = (Mycells0.children[1] == null) ? Mycells0.children[0] : Mycells0.children[1];
                }
                else {
                    MyCheck = Mycells0.children[2];
                }
                if (MyCheck.checked) {
                    flag = true;
                    if (OCID != '') OCID += ',';
                    OCID += '\'' + MyCheck.value + '\'';
                }
            }

            document.getElementById('PrintValue').value = OCID;


            if (document.getElementById('PrintValue').value == '') {
                alert('請勾選要列印的學員班級!')
                return false;
            }
            else {
                //alert(document.getElementById('PrintValue').value);
                uUrl = '../../SQControl.aspx?&SQ_AutoLogout=true&sys=Member&filename=SD_09_015_R2'
                    + '&OCID=' + document.getElementById('PrintValue').value + '&RID=' + document.getElementById('RIDValue').value;
                if (type == 'Y') {
                    uUrl = '../../SQControl.aspx?&SQ_AutoLogout=true&sys=Member&filename=SD_09_015_R1'
                        + '&OCID=' + document.getElementById('PrintValue').value + '&RID=' + document.getElementById('RIDValue').value;
                }
                window.open(uUrl, 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
            }
        }

        function InsertValue(Flag, MyValue) {
            //alert(Flag);
            //alert(MyValue);
            if (Flag) {
                if (document.getElementById('PrintValue').value.indexOf('\'' + MyValue + '\'') == -1) {
                    if (document.getElementById('PrintValue').value == '') {
                        document.getElementById('PrintValue').value = '\'' + MyValue + '\''
                    }
                    else {
                        document.getElementById('PrintValue').value += ',\'' + MyValue + '\''
                    }
                }
            }
            else {
                if (document.getElementById('PrintValue').value.indexOf('\'' + MyValue + '\'') != -1) {
                    if (document.getElementById('PrintValue').value.indexOf(',\'' + MyValue + '\'') != -1) {
                        document.getElementById('PrintValue').value = document.getElementById('PrintValue').value.replace(',\'' + MyValue + '\'', '');
                    }
                    if (document.getElementById('PrintValue').value.indexOf('\'' + MyValue + '\',') != -1) {
                        document.getElementById('PrintValue').value = document.getElementById('PrintValue').value.replace('\'' + MyValue + '\',', '');
                    }
                    if (document.getElementById('PrintValue').value.indexOf('\'' + MyValue + '\'') != -1) {
                        document.getElementById('PrintValue').value = document.getElementById('PrintValue').value.replace('\'' + MyValue + '\'', '');
                    }
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;<font color="#990000">補助/未補助學員名冊</font>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="100">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button12" Style="display: none" runat="server" Text="Button12"></asp:Button>
                                <span id="HistoryList2" style="display: none; z-index: 100; position: absolute" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="100">班別/職類
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="205px"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="205px"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <span id="HistoryList" style="display: none; z-index: 101; left: 270px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="100">期別
                            </td>
                            <td class="whitecol">
                                <input id="CyclType" type="text" style="width: 32px; height: 22px" runat="server" name="CyclType" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="100">班級範圍
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="ClassRound" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="開訓二週前">開訓二週前</asp:ListItem>
                                    <asp:ListItem Value="已開訓">已開訓</asp:ListItem>
                                    <asp:ListItem Value="已結訓">已結訓</asp:ListItem>
                                    <asp:ListItem Value="未開訓">未開訓</asp:ListItem>
                                    <asp:ListItem Value="全部" Selected="True">全部</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <p align="center">
                                    <asp:Button ID="Query" runat="server" Text="查詢" CssClass="button_b_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <p align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td style="font-size: 10pt; color: red">
                                <font face="新細明體">(有 * 號的班級代表至少有一個學員，於學員資料維護的必填資料未填)</font>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="選取">
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                            <HeaderTemplate>
                                                選取<input id="CheckboxAll" type="checkbox" runat="server" name="CheckboxAll" />
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="LabelStar" runat="server" ForeColor="Red">*</asp:Label>
                                                <input id="Checkbox1" type="checkbox" value='<%# DataBinder.Eval(Container.DataItem,"OCID")%>' runat="server" name="Checkbox1" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassID" HeaderText="班級代碼">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="IsApplic" HeaderText="志願班別">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                    <asp:Button ID="Button1" runat="server" Text="列印補助學員名冊" CssClass="asp_Export_M"></asp:Button>
                                    <asp:Button ID="Button3" runat="server" Text="列印未補助學員名冊" CssClass="asp_Export_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="PrintValue" type="hidden" runat="server" name="PrintValue">
    </form>
</body>
</html>
