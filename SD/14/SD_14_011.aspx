<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_011.aspx.vb" Inherits="WDAIIP.SD_14_011" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>預估參訓學員補助經費清冊</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //function printkind() {
        //    var print_orderyby = document.getElementById('print_orderyby');
        //    var Hidorder1 = document.getElementById('Hidorder1');
        //    var Hidorder2 = document.getElementById('Hidorder2');
        //    print_orderyby.value = Hidorder1.value; //'R.IDNO'
        //    if (getValue("print_type") == '2') {
        //        print_orderyby.value = Hidorder2.value; //'R.StudentID'
        //    }
        //}

        function CheckSearch() {
            var vSTDate1 = document.getElementById('STDate1').value;
            var vSTDate2 = document.getElementById('STDate2').value;
            var vFTDate1 = document.getElementById('FTDate1').value;
            var vFTDate2 = document.getElementById('FTDate2').value;
            var msg = '';
            if (!checkDate(vSTDate1) && vSTDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
            if (!checkDate(vSTDate2) && vSTDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';
            if (!checkDate(vFTDate1) && vFTDate1 != '') msg += '結訓起始日期必須為正確日期格式\n';
            if (!checkDate(vFTDate2) && vFTDate2 != '') msg += '結訓結束日期必須為正確日期格式\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows[i].cells[0].children[0].checked = Flag;
                SelectItem(Flag, MyTable.rows[i].cells[0].children[0].value);
            }
        }

        function SelectItem(Flag, MyValue) {
            var SelectValue = document.getElementById('SelectValue');
            if (Flag) {
                //add
                if (SelectValue.value != '') { SelectValue.value += ',' }
                SelectValue.value += MyValue;
            }
            else {
                //delete
                if (SelectValue.value.indexOf(',' + MyValue) != -1)
                    SelectValue.value = SelectValue.value.replace(',' + MyValue, '')
                else if (SelectValue.value.indexOf(MyValue + ',') != -1)
                    SelectValue.value = SelectValue.value.replace(MyValue + ',', '')
                else if (SelectValue.value.indexOf(MyValue) != -1)
                    SelectValue.value = SelectValue.value.replace(MyValue, '')
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;預估參訓學員補助經費清冊</asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓期間</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                                <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓期間</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                                <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <input id="Button4" type="button" value="列印機構所有班級" name="Button4" runat="server" class="asp_button_M" />
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" HorizontalAlign="Center" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <HeaderTemplate>
                                                <input type="checkbox" onclick="SelectAll(this.checked);">
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <input id="OCID" type="checkbox" runat="server" name="OCID">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="23%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                            <HeaderStyle Width="23%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="23%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="23%"></HeaderStyle>
                                        </asp:BoundColumn>
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
                        <tr>
                            <td align="center">
                                <asp:RadioButtonList ID="print_type" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">依【身分證字號】排序列印&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</asp:ListItem>
                                    <asp:ListItem Value="2" Selected="True">依【學號】排序列印</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button5" runat="server" Text="列印" CssClass="asp_Export_M" />&nbsp;</td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="Years" type="hidden" name="Years" runat="server">
        <input id="SelectValue" type="hidden" name="SelectValue" runat="server">
        <input id="PlanID" type="hidden" name="PlanID" runat="server">
        <input id="orgid" type="hidden" name="orgid" runat="server">
        <input id="print_orderyby" type="hidden" runat="server">
        <%--<input id="Hidorder1" type="hidden" runat="server">
        <input id="Hidorder2" type="hidden" runat="server">--%>
    </form>
</body>
</html>
