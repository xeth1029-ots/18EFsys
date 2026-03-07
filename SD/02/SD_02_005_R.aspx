<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_02_005_R.aspx.vb" Inherits="WDAIIP.SD_02_005_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>甄試成績表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }

        /*
		function choose_class(){				
		openClass('SD_02_ch.aspx');
		}
		function ReportPrint(){
		if (document.form1.OCIDValue1.value==''){
		alert('請選擇職類');
		return false;
		}
		return true;
		}
		*/

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }

        function showPanel() {
            //debugger;
            document.getElementById('tr1').style.display = '';
            //document.getElementById('tr1').style.display = 'inline';
            document.getElementById('tr2').style.display = '';
            //document.getElementById('tr2').style.display = 'inline';
            document.getElementById('tr3').style.display = 'none';
            document.getElementById('tr4').style.display = 'none';
            if (document.form1.radiobtn1[1].checked) {
                document.getElementById('tr1').style.display = 'none';
                document.getElementById('tr2').style.display = 'none';
                document.getElementById('tr3').style.display = '';
                document.getElementById('tr4').style.display = '';
                //document.getElementById('tr3').style.display = 'inline';
                //document.getElementById('tr4').style.display = 'inline';
            }
        }

        function CheckData() {
            //debugger;
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var trainValue = document.getElementById('trainValue');
            var STDate1 = document.getElementById('STDate1');
            var STDate2 = document.getElementById('STDate2');
            var msg = '';
            if (document.form1.radiobtn1[1].checked) {
                if (trainValue.value == '') msg += '請選擇訓練職類\n';
                if (STDate1.value == '') msg += '開訓日期的起始日不能是空白\n';
                if (STDate2.value == '') msg += '開訓日期的結束日不能是空白\n';
                if (!IsDate(STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
                if (!IsDate(STDate2.value)) msg += '開訓日期的結束日不是正確的日期格式\n';
            }
            else {
                if (OCIDValue1.value == '') msg += '請選擇訓練機構／班別\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function IsDate(MyDate) {
            if (MyDate != '') {
                if (!checkDate(MyDate))
                    return false;
            }
            return true;
        }

        function ChangeAll(obj) {
            var objLen = document.form1.length;
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (obj.checked == true) {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        document.form1.elements[iCount].checked = true;
                    }
                }
                else {
                    if (document.form1.elements[iCount].type == "checkbox") {
                        document.form1.elements[iCount].checked = false;
                    }
                }
            }
        }

        function CheckPrint() {
            //var flag = false;
            var MyTable = document.getElementById('DataGrid1');
            var OCID = '';
            //var OCID=document.getElementById('OCIDValue1').value;
            for (var i = 1; i < MyTable.rows.length; i++) {
                var MyCheck = MyTable.rows[i].cells[0].children[0];
                if (MyCheck.checked) {
                    //flag = true;
                    if (OCID != '') { OCID += ','; }
                    OCID += '\'' + MyCheck.value + '\'';
                }
            }
            var PrintValue = document.getElementById('PrintValue');
            PrintValue.value = OCID;
            if (PrintValue.value == '') {
                alert('請勾選本頁要列印的學員班級!');
                return false;
            }
            //else {
            //    alert(document.getElementById('PrintValue').value);
            //    url = '../../SQControl.aspx?&SQ_AutoLogout=true&sys=Member&filename=Maintest_GradeList&path=TIMS&'
            //    window.open(url + 'OCID1=' + document.getElementById('PrintValue').value, 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
            //}
        }

        function InsertValue(Flag, MyValue) {
            //alert(Flag);
            //alert(MyValue);
            var PrintValue = document.getElementById('PrintValue');
            var xMyValue = '\'' + MyValue + '\'';
            if (Flag) {
                if (PrintValue.value.indexOf(xMyValue) == -1) {
                    if (PrintValue.value != '') { PrintValue.value += ','; }
                    PrintValue.value += xMyValue;
                }
            }
            else {
                if (PrintValue.value.indexOf(xMyValue) != -1) {
                    if (PrintValue.value.indexOf(',' + xMyValue) != -1) {
                        PrintValue.value = PrintValue.value.replace(',' + xMyValue, '');
                    }
                    if (PrintValue.value.indexOf(xMyValue + ',') != -1) {
                        PrintValue.value = PrintValue.value.replace(xMyValue + ',', '');
                    }
                    if (PrintValue.value.indexOf(xMyValue) != -1) {
                        PrintValue.value = PrintValue.value.replace(xMyValue, '');
                    }
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="titlelab1" runat="server"></asp:Label>
                    <asp:Label ID="titlelab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;甄試成績表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" id="table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">依據選擇</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="radiobtn1" runat="server" RepeatLayout="flow" RepeatDirection="horizontal" AutoPostBack="true" CssClass="font">
                                    <asp:ListItem Value="1" Selected="true">依訓練機構／班別</asp:ListItem>
                                    <asp:ListItem Value="2">依訓練職類／開訓期間</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr1" runat="server">
                            <td class="bluecol">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="button8" type="button" value="..." name="button5" runat="server" class="asp_button_Mini" />
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3" CssClass="asp_button_S"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="historyrid" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="tr2" runat="server">
                            <td class="bluecol_need">職類/班別</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="historytable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="tr3" runat="server">
                            <td class="bluecol_need">訓練職類</td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." runat="server" class="asp_button_Mini" />
                                <input id="trainValue" type="hidden" runat="server" />
                            </td>
                        </tr>
                        <tr id="tr4" runat="server">
                            <td class="bluecol_need">開訓期間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                ～
                                <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                        <tr id="Trwork2013a" runat="server">
                            <td class="bluecol">就服單位協助報名</td>
                            <td class="whitecol">
                                <asp:RadioButtonList Style="z-index: 0" ID="rblEnterPathW" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="Y">是</asp:ListItem>
                                    <asp:ListItem Value="N">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Button ID="query" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button><br />
                        <asp:Label ID="msg" runat="server" ForeColor="red"></asp:Label>
                    </div>
                    <table id="table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="false" AllowPaging="true" CssClass="font" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="選取">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderTemplate>選取<input id="checkboxall" type="checkbox" runat="server"></HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="checkbox1" type="checkbox" value='<%# databinder.eval(container.dataitem,"ocid")%>' runat="server"></ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="classcname" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="stdate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ftdate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="false"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <div align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="PrintValue" type="hidden" runat="server">
    </form>
</body>
</html>
