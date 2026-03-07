<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_01_001_add.aspx.vb" Inherits="WDAIIP.CM_01_001_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練計畫核銷作業</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script src="../../js/date-picker.js"></script>
    <script src="../../js/common.js"></script>
    <script src="../../js/OpenWin/openwin.js"></script>
    <script>
        function CheckData1() {
            var msg = '';
            if (document.getElementById('CancelDate1').value == '') msg += '請輸入送核日期\n';
            else if (!checkDate(document.getElementById('CancelDate1').value)) msg += '送核日期必須是正確的日期格式\n';
            if (document.getElementById('PCID').selectedIndex == 0) msg += '請選擇核銷項目\n';
            if (document.getElementById('Itemage').value == '') msg += '請輸入數量\n';
            else if (!isUnsignedInt(document.getElementById('Itemage').value)) msg += '數量必須為整數\n';
            if (document.getElementById('ItemCost').value == '') msg += '請輸入計價單位\n';
            else if (!isUnsignedInt(document.getElementById('ItemCost').value)) msg += '計價單位必須為整數\n';
            if (!isChecked(document.getElementsByName('PlanType1'))) msg += '請選擇計畫分類\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        function CheckData2() {
            var MyTable = document.getElementById('DataGrid2');
            var msg = '';
            if (document.getElementById('CancelDate2').value == '') msg += '請輸入送核日期\n';
            else if (!checkDate(document.getElementById('CancelDate2').value)) msg += '送核日期必須是正確的日期格式\n';
            if (document.getElementById('Times1').value == '') msg += '請輸入核銷次數\n';
            else if (!isUnsignedInt(document.getElementById('Times1').value)) msg += '核銷次數必須為整數\n';
            else {
                if (MyTable) {
                    for (i = 1; i < MyTable.rows.length; i++) {
                        if (parseInt(document.getElementById('Times1').value) == parseInt(MyTable.rows(i).cells(0).innerHTML)) msg += '不能輸入相同的核銷次數\n'
                    }
                }
            }
            if (document.getElementById('PNum').value == '') msg += '請輸入核銷人數\n';
            else if (!isUnsignedInt(document.getElementById('PNum').value)) msg += '核銷人數必須為整數\n';
            if (document.getElementById('PMoney').value == '') msg += '請輸入核銷金額\n';
            else if (!isDesignFloat(document.getElementById('PMoney').value)) msg += '核銷金額最小只能到小數點第二位\n';
            if (!isChecked(document.getElementsByName('BudID'))) msg += '請選擇預算種類\n';
            if (!isChecked(document.getElementsByName('PlanType2'))) msg += '請選擇計畫分類\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        function CheckData3() {
            var MyTable = document.getElementById('DataGrid3');
            var Total = 0;
            var msg = '';
            if (document.getElementById('CancelDate3').value == '') msg += '請輸入送核日期\n';
            else if (!checkDate(document.getElementById('CancelDate3').value)) msg += '送核日期必須是正確的日期格式\n';
            if (document.getElementById('Times2').value == '') msg += '請輸入核銷次數\n';
            else if (!isUnsignedInt(document.getElementById('Times2').value)) msg += '核銷次數必須為整數\n';
            else {
                if (MyTable) {
                    for (i = 1; i < MyTable.rows.length; i++) {
                        if (parseInt(document.getElementById('Times2').value) == parseInt(MyTable.rows(i).cells(0).innerHTML)) msg += '不能輸入相同的核銷次數\n'
                    }
                }
            }
            if (document.getElementById('SelfPrice').value == '') msg += '請輸入個人訓練費用單價\n';
            else if (!isDesignFloat(document.getElementById('SelfPrice').value)) msg += '個人訓練費用單價最小只能到小數點第二位\n';

            if (document.getElementById('Num3').value != '' && !isUnsignedInt(document.getElementById('Num3').value)) msg += '就安公費100%[人數]必須為整數\n';
            if (document.getElementById('Percent3').value != '') {
                if (!isUnsignedInt(document.getElementById('Percent3').value)) msg += '就安公費100%[百分比]必須為整數\n';
                Total = 100;
                if (MyTable) {
                    for (i = 1; i < MyTable.rows.length; i++) {
                        Total = Total - MyTable.rows(i).cells(9).children(3).value;
                    }
                }
                if (Total < parseInt(document.getElementById('Percent3').value)) msg += '就安公費不能超過100%(剩餘' + Total + '%)\n'
            }
            if (document.getElementById('Num4').value != '') {
                if (!isUnsignedInt(document.getElementById('Num4').value)) msg += '就安自費80%[人數]必須為整數\n';
                Total = 80;
                if (MyTable) {
                    for (i = 1; i < MyTable.rows.length; i++) {
                        Total = Total - MyTable.rows(i).cells(9).children(4).value;
                    }
                }
                if (Total < parseInt(document.getElementById('Percent4').value)) msg += '就安自費不能超過80%(剩餘' + Total + '%)\n'
            }
            if (document.getElementById('Num1').value != '' && !isUnsignedInt(document.getElementById('Num1').value)) msg += '就保公費100%[人數]必須為整數\n';
            if (document.getElementById('Percent1').value != '') {
                if (!isUnsignedInt(document.getElementById('Percent1').value)) msg += '就保公費100%[百分比]必須為整數\n';
                Total = 100;
                if (MyTable) {
                    for (i = 1; i < MyTable.rows.length; i++) {
                        Total = Total - MyTable.rows(i).cells(9).children(1).value;
                    }
                }
                if (Total < parseInt(document.getElementById('Percent1').value)) msg += '就保公費不能超過100%(剩餘' + Total + '%)\n'
            }
            if (document.getElementById('Num2').value != '' && !isUnsignedInt(document.getElementById('Num2').value)) msg += '就保自費80%[人數]必須為整數\n';
            if (document.getElementById('Percent2').value != '') {
                if (!isUnsignedInt(document.getElementById('Percent2').value)) msg += '就保自費80%[百分比]必須為整數\n';
                Total = 80;
                if (MyTable) {
                    for (i = 1; i < MyTable.rows.length; i++) {
                        Total = Total - MyTable.rows(i).cells(9).children(2).value;
                    }
                }
                if (Total < parseInt(document.getElementById('Percent2').value)) msg += '就保自費不能超過80%(剩餘' + Total + '%)\n'
            }
            if (document.getElementById('Percent4').value != '' && !isUnsignedInt(document.getElementById('Percent4').value)) msg += '就安自費80%[百分比]必須為整數\n';
            if (!isChecked(document.getElementsByName('PlanType3'))) msg += '請選擇計畫分類\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        function CheckData4() {
            var MyTable = document.getElementById('DataGrid4');
            var msg = '';
            if (document.getElementById('CancelDate4').value == '') msg += '請輸入送核日期\n';
            else if (!checkDate(document.getElementById('CancelDate4').value)) msg += '送核日期必須是正確的日期格式\n';
            if (document.getElementById('Times3').value == '') msg += '請輸入核銷次數\n';
            else if (!isUnsignedInt(document.getElementById('Times3').value)) msg += '核銷次數必須為整數\n';
            else {
                if (MyTable) {
                    for (i = 1; i < MyTable.rows.length; i++) {
                        if (parseInt(document.getElementById('Times3').value) == parseInt(MyTable.rows(i).cells(0).innerHTML)) msg += '不能輸入相同的核銷次數\n'
                    }
                }
            }
            if (!isChecked(document.getElementsByName('PlanType4'))) msg += '請選擇計畫分類\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function CheckData5() {
            var MyTable = document.getElementById('DataGrid5');
            var msg = '';
            if (document.getElementById('CancelDate5').value == '') msg += '請輸入送核日期\n';
            else if (!checkDate(document.getElementById('CancelDate5').value)) msg += '送核日期必須是正確的日期格式\n';
            if (document.getElementById('Times4').value == '') msg += '請輸入核銷次數\n';
            else if (!isUnsignedInt(document.getElementById('Times4').value)) msg += '核銷次數必須為整數\n';
            else {
                if (MyTable) {
                    for (i = 1; i < MyTable.rows.length; i++) {
                        if (parseInt(document.getElementById('Times4').value) == parseInt(MyTable.rows(i).cells(0).innerHTML)) msg += '不能輸入相同的核銷次數\n'
                    }
                }
            }
            if (document.getElementById('GNum').value != '' && !isUnsignedInt(document.getElementById('GNum').value)) msg += '就保公費100%[人數]必須為整數\n';
            if (document.getElementById('GPrice').value != '' && !isDesignFloat(document.getElementById('GPrice').value)) msg += '平均單價最小只能到小數點第二位\n';
            if (document.getElementById('SNum').value != '' && !isUnsignedInt(document.getElementById('SNum').value)) msg += '就保公費100%[人數]必須為整數\n';
            if (document.getElementById('SPrice').value != '' && !isDesignFloat(document.getElementById('SPrice').value)) msg += '平均單價最小只能到小數點第二位\n';
            if (!isChecked(document.getElementsByName('PlanType5'))) msg += '請選擇計畫分類\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function isDesignFloat(value) {
            if (isUnsignedInt(value)) return true;
            else {
                if (!isPositiveFloat(value)) return false;
                else {
                    var pattern = /^\d+\.{0,1}\d{0,2}$/;
                    return pattern.test(value);
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
                        <td class="font">
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                            <asp:Label ID="TitleLab2" runat="server">
				            首頁&gt;&gt;訓練經費控管&gt;&gt;<font color="#990000">訓練計畫核銷作業</font>
                            </asp:Label>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="Table2">
                    <tr>
                        <td width="100" class="bluecol">
                            單位名稱
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="OrgName" runat="server"></asp:Label>
                        </td>
                        <td width="100" class="bluecol">
                        </td>
                        <td class="whitecol">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            班別名稱
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="ClassCName" runat="server"></asp:Label>
                        </td>
                        <td class="bluecol">
                            開結訓日期：
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="TDate" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            訓練經費
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="TrainCost" runat="server"></asp:Label>
                        </td>
                        <td class="bluecol">
                            結餘總金額：
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="CancelCost" runat="server"></asp:Label>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            開訓人數
                        </td>
                        <td class="whitecol">
                            <asp:Label ID="Tnum" runat="server"></asp:Label>
                        </td>
                        <td class="bluecol">
                        </td>
                        <td class="whitecol">
                        </td>
                    </tr>
                </table>
                <table id="CancelMode1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    <tr>
                        <td>
                            <table class="table_sch" id="Table3">
                                <tr>
                                    <td class="bluecol_need" width="100">
                                        送核日期
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="CancelDate1" runat="server" Columns="10"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= CancelDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        經費項目
                                    </td>
                                    <td class="whitecol">
                                        <font face="新細明體">
                                            <asp:DropDownList ID="PCID" runat="server">
                                            </asp:DropDownList>
                                        </font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        數量/計價單位
                                    </td>
                                    <td class="whitecol">
                                        <font face="新細明體">
                                            <asp:TextBox ID="Itemage" runat="server" Columns="7"></asp:TextBox>/
                                            <asp:TextBox ID="ItemCost" runat="server" Columns="7"></asp:TextBox></font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        計劃分類
                                    </td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="PlanType1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="1">自辦</asp:ListItem>
                                            <asp:ListItem Value="2">委辦</asp:ListItem>
                                            <asp:ListItem Value="3">合辦</asp:ListItem>
                                            <asp:ListItem Value="4">補助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        備註說明
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Note1" runat="server" Columns="50" TextMode="MultiLine" Rows="5"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" colspan="2" class="whitecol">
                                        <asp:Button ID="Button1" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="Button3" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button>
                                    </td>
                                </tr>
                            </table>
                            <table id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td>
                                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AutoGenerateColumns="False" Width="100%">
                                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <Columns>
                                                <asp:BoundColumn HeaderText="序號">
                                                    <HeaderStyle Width="25px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelDate" HeaderText="送核日期" DataFormatString="{0:d}">
                                                    <HeaderStyle Width="60px"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CostName" HeaderText="經費項目"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="BudgetCost" HeaderText="預算金額" DataFormatString="{0:#,##0.00}">
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="AddUpCost" HeaderText="已送核&lt;BR&gt;累積金額" DataFormatString="{0:#,##0.00}">
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Itemage" HeaderText="數量">
                                                    <HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="ItemCost" HeaderText="計價單位">
                                                    <HeaderStyle Width="60px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelCost" HeaderText="送核金額" DataFormatString="{0:#,##0.00}">
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn HeaderText="結餘金額" DataFormatString="{0:#,##0.00}">
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Note" HeaderText="備註">
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="50px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Button ID="Button2" runat="server" Text="刪除" CommandName="delete2"></asp:Button>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <table id="CancelMode2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    <tr>
                        <td>
                            <table class="table_sch" id="Table4">
                                <tr>
                                    <td class="bluecol_need" width="100">
                                        <font face="新細明體">&nbsp;&nbsp;&nbsp; 送核日期<font color="red">*</font></font>
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="CancelDate2" runat="server" Columns="10"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= CancelDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        核銷次數
                                    </td>
                                    <td class="whitecol">
                                        <font face="新細明體">
                                            <asp:TextBox ID="Times1" runat="server" Columns="7"></asp:TextBox></font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        核銷人數
                                    </td>
                                    <td class="whitecol">
                                        <font face="新細明體">
                                            <asp:TextBox ID="PNum" runat="server" Columns="7"></asp:TextBox></font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        核銷金額
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="PMoney" runat="server" Columns="10"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        預算種類
                                    </td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="BudID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        計劃分類
                                    </td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="PlanType2" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="1">自辦</asp:ListItem>
                                            <asp:ListItem Value="2">委辦</asp:ListItem>
                                            <asp:ListItem Value="3">合辦</asp:ListItem>
                                            <asp:ListItem Value="4">補助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        核銷說明
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Note2" runat="server" Columns="50" TextMode="MultiLine" Rows="5"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" colspan="2" class="whitecol">
                                        <asp:Button ID="Button5" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="Button4" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button>
                                    </td>
                                </tr>
                            </table>
                            <table id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td>
                                        <asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" AutoGenerateColumns="False" Width="100%">
                                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <Columns>
                                                <asp:BoundColumn DataField="Times" HeaderText="次數">
                                                    <HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelDate" HeaderText="送核日期" DataFormatString="{0:d}">
                                                    <HeaderStyle Width="60px"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="BudName" HeaderText="預算種類"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="PNum" HeaderText="核銷人數"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="PMoney" HeaderText="金額" DataFormatString="{0:#,##0.00}">
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelCost" HeaderText="合計金額" DataFormatString="{0:#,##0.00}">
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Note" HeaderText="核銷說明"></asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="50px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Button ID="Button6" runat="server" Text="刪除" CommandName="delete6"></asp:Button>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <table class="font" id="CancelMode3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    <tr>
                        <td>
                            <table class="table_sch" id="Table5">
                                <tr>
                                    <td class="bluecol_need" width="100">
                                        送核日期
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="CancelDate3" runat="server" Columns="10"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= CancelDate3.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        核銷次數
                                    </td>
                                    <td class="whitecol">
                                        <font face="新細明體">
                                            <asp:TextBox ID="Times2" runat="server" Columns="7"></asp:TextBox></font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        個人訓練費用<br>
                                        單價
                                    </td>
                                    <td class="whitecol">
                                        <font face="新細明體">
                                            <asp:TextBox ID="SelfPrice" runat="server" Columns="7"></asp:TextBox></font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        就安公費100%
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Num3" runat="server" Columns="3"></asp:TextBox>人*個人訓練費用單價*
                                        <asp:TextBox ID="Percent3" runat="server" Columns="3"></asp:TextBox>%
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        就安自費80%
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Num4" runat="server" Columns="3"></asp:TextBox>人*個人訓練費用單價*
                                        <asp:TextBox ID="Percent4" runat="server" Columns="3"></asp:TextBox>%
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        就保公費100%
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Num1" runat="server" Columns="3"></asp:TextBox>人*個人訓練費用單價*
                                        <asp:TextBox ID="Percent1" runat="server" Columns="3"></asp:TextBox>%
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        就保自費80%
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Num2" runat="server" Columns="3"></asp:TextBox>人*個人訓練費用單價*
                                        <asp:TextBox ID="Percent2" runat="server" Columns="3"></asp:TextBox>%
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        計劃分類
                                    </td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="PlanType3" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="1">自辦</asp:ListItem>
                                            <asp:ListItem Value="2">委辦</asp:ListItem>
                                            <asp:ListItem Value="3">合辦</asp:ListItem>
                                            <asp:ListItem Value="4">補助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        核銷說明
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Note3" runat="server" Columns="50" TextMode="MultiLine" Rows="5"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" colspan="2" class="whitecol">
                                        <asp:Button ID="Button7" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="Button8" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button>&nbsp;<input id="Button10" style="width: 150px" type="button" value="調整班級學員的預算別" name="Button10" runat="server" class="asp_button_L">
                                    </td>
                                </tr>
                            </table>
                            <table id="DataGridTable3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td>
                                        <asp:DataGrid ID="DataGrid3" runat="server" CssClass="font" AutoGenerateColumns="False" Width="100%">
                                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <Columns>
                                                <asp:BoundColumn DataField="Times" HeaderText="次數">
                                                    <HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelDate" HeaderText="送核日期" DataFormatString="{0:d}">
                                                    <HeaderStyle Width="60px"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="SelfPrice" HeaderText="平均單價" DataFormatString="{0:#,##0.00}">
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn HeaderText="就安公費"></asp:BoundColumn>
                                                <asp:BoundColumn HeaderText="就安自費"></asp:BoundColumn>
                                                <asp:BoundColumn HeaderText="就保公費"></asp:BoundColumn>
                                                <asp:BoundColumn HeaderText="就保自費"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelCost" HeaderText="送核金額小計" DataFormatString="{0:#,##0.00}">
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Note" HeaderText="核銷說明"></asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="50px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Button ID="Button9" runat="server" Text="刪除"></asp:Button><input id="PValue1" type="hidden" runat="server"><input id="PValue2" type="hidden" runat="server"><input id="PValue3" type="hidden" runat="server"><input id="PValue4" type="hidden" runat="server">
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <table class="font" id="CancelMode4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    <tr>
                        <td>
                            <table class="font" id="Table6" width="100%" cellpadding="1" cellspacing="1">
                                <tr>
                                    <td class="bluecol_need" width="100">
                                        送核日期
                                    </td>
                                    <td class="td_light">
                                        <asp:TextBox ID="CancelDate4" runat="server" Columns="10"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= CancelDate3.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        核銷次數
                                    </td>
                                    <td class="td_light">
                                        <font face="新細明體">
                                            <asp:TextBox ID="Times3" runat="server" Columns="7"></asp:TextBox></font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        計劃分類
                                    </td>
                                    <td class="td_light">
                                        <asp:RadioButtonList ID="PlanType4" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="1">自辦</asp:ListItem>
                                            <asp:ListItem Value="2">委辦</asp:ListItem>
                                            <asp:ListItem Value="3">合辦</asp:ListItem>
                                            <asp:ListItem Value="4">補助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        核銷說明
                                    </td>
                                    <td class="td_light">
                                        <asp:TextBox ID="Note4" runat="server" Columns="50" TextMode="MultiLine" Rows="5"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        <asp:Table ID="DGTable" runat="server" CssClass="font" Width="500px" CellPadding="3" CellSpacing="0">
                                        </asp:Table>
                                        <asp:Button ID="Button18" runat="server" Text="產生名單(隱藏)" CssClass="asp_button_M"></asp:Button>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" colspan="2">
                                        <asp:Button ID="Button15" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button><asp:Button ID="Button16" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button>
                                        <input id="Button17" style="width: 160px" type="button" value="調整班級學員的參加單元" name="Button10" runat="server" class="button_b_L">
                                    </td>
                                </tr>
                            </table>
                            <table id="DataGridTable4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td>
                                        <asp:DataGrid ID="DataGrid4" runat="server" CssClass="font" AutoGenerateColumns="False" Width="100%">
                                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <Columns>
                                                <asp:BoundColumn DataField="Times" HeaderText="次數">
                                                    <HeaderStyle Width="25px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelDate" HeaderText="送核日期" DataFormatString="{0:d}">
                                                    <HeaderStyle Width="60px"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="TPrice1" HeaderText="就安" DataFormatString="{0:#,##0.00}">
                                                    <HeaderStyle Width="100px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="TPrice2" HeaderText="就保" DataFormatString="{0:#,##0.00}">
                                                    <HeaderStyle Width="100px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelCost" HeaderText="送核金額小計" DataFormatString="{0:#,##0.00}">
                                                    <HeaderStyle Width="100px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="Note" HeaderText="核銷說明"></asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="50px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Button ID="Button19" runat="server" Text="刪除"></asp:Button>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
                <table class="font" id="CancelMode5" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    <tr>
                        <td>
                            <table class="table_sch" id="Table6">
                                <tr>
                                    <td class="bluecol_need" width="100">
                                        送核日期
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="CancelDate5" runat="server" Columns="10"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= CancelDate3.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        核銷次數
                                    </td>
                                    <td class="whitecol">
                                        <font face="新細明體">
                                            <asp:TextBox ID="Times4" runat="server" Columns="7"></asp:TextBox></font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        一般對象
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="GNum" runat="server" Columns="3"></asp:TextBox>人*平均單價
                                        <asp:TextBox ID="GPrice" runat="server" Columns="5"></asp:TextBox>元
                                        <asp:Label ID="Var1" runat="server"></asp:Label><input id="ItemVar1" type="hidden" runat="server">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        特定對象
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="SNum" runat="server" Columns="3"></asp:TextBox>人*平均單價
                                        <asp:TextBox ID="SPrice" runat="server" Columns="5"></asp:TextBox>元
                                        <asp:Label ID="Var2" runat="server"></asp:Label><input id="ItemVar2" type="hidden" runat="server">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need">
                                        計劃分類
                                    </td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="PlanType5" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                            <asp:ListItem Value="1">自辦</asp:ListItem>
                                            <asp:ListItem Value="2">委辦</asp:ListItem>
                                            <asp:ListItem Value="3">合辦</asp:ListItem>
                                            <asp:ListItem Value="4">補助</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">
                                        核銷說明
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="Note5" runat="server" Columns="50" TextMode="MultiLine" Rows="5"></asp:TextBox>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center" colspan="2" class="whitecol">
                                        <asp:Button ID="Button11" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="Button12" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button>&nbsp;<input id="Button13" style="width: 190px" type="button" value="調整學員的主要參訓身分別" name="Button10" runat="server" class="button_b_L">
                                    </td>
                                </tr>
                            </table>
                            <table id="DataGridTable5" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td>
                                        <asp:DataGrid ID="DataGrid5" runat="server" CssClass="font" AutoGenerateColumns="False" Width="100%">
                                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <Columns>
                                                <asp:BoundColumn DataField="Times" HeaderText="次數">
                                                    <HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelDate" HeaderText="送核日期" DataFormatString="{0:d}">
                                                    <HeaderStyle Width="60px"></HeaderStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn HeaderText="一般身分"></asp:BoundColumn>
                                                <asp:BoundColumn HeaderText="特地對象"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="CancelCost" HeaderText="送核金額小計" DataFormatString="{0:#,##0.00}"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="Note" HeaderText="核銷說明"></asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="50px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Button ID="Button14" runat="server" Text="刪除"></asp:Button>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <input id="CancelModeValue" type="hidden" runat="server">
    </form>
</body>
</html>
