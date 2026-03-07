<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_010_R.aspx.vb" Inherits="WDAIIP.SD_11_010_R" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>受訓期間滿意度調查統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function chk() {
            var msg = '';
            if (document.form1.center.value == '') msg += '請選擇訓練機構\n';
            if (document.form1.STDate1.value != '') {
                if (!IsDate(document.form1.STDate1.value)) msg += '開訓日期的起始日 不是正確的日期格式\n';
            }
            if (document.form1.STDate2.value != '') {
                if (!IsDate(document.form1.STDate2.value)) msg += '開訓日期的迄日 不是正確的日期格式\n';
            }
            if (document.form1.FTDate1.value != '') {
                if (!IsDate(document.form1.FTDate1.value)) msg += '結訓日期的起始日 不是正確的日期格式\n';
            }
            if (document.form1.FTDate2.value != '') {
                if (!IsDate(document.form1.FTDate2.value)) msg += '結訓日期的迄日 不是正確的日期格式\n';
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

        //function ClearData() {
        //    document.getElementById('PlanID').value = '';
        //    document.getElementById('center').value = '';
        //    document.getElementById('RIDValue').value = '';
        //    for (var i = document.form1.OCID.options.length - 1; i >= 0; i--) {
        //        document.form1.OCID.options[i] = null;
        //    }
        //    document.getElementById('OCID').style.display = 'none';
        //    //document.getElementById('msg').innerHTML = '請先選擇機構';
        //}


        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
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
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;訓練成效與滿意度&gt;&gt;受訓期間滿意度調查統計表</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="myTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1" width="100%">
                        <tr id="Year_TR" runat="server">
                            <td class="bluecol" style="width: 20%">年度</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="yearlist" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr id="DistID_TR" runat="server">
                            <td class="bluecol">轄區</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="DistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server" size="1">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">辦訓地縣市</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="Tcitycode" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="7"></asp:CheckBoxList>
                                <input id="TcityHidden" type="hidden" value="0" name="TcityHidden" runat="server" size="1">
                            </td>
                        </tr>
                        <%--<tr id="TPlanID0_TR" runat="server">
                            <td class="bluecol">訓練計畫(職前)</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chkTPlanID0" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0"></asp:CheckBoxList>
                                <input id="TPlanID0HID" type="hidden" value="0" name="TPlanID0HID" runat="server" size="1">
                            </td>
                        </tr>
                        <tr id="TPlanID1_TR" runat="server">
                            <td class="bluecol">訓練計畫(在職)</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chkTPlanID1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0"></asp:CheckBoxList>
                                <input id="TPlanID1HID" type="hidden" value="0" name="TPlanID1HID" runat="server" size="1">
                            </td>
                        </tr>
                        <tr id="TPlanIDX_TR" runat="server">
                            <td class="bluecol">訓練計畫(其他)</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chkTPlanIDX" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3" CellPadding="0" CellSpacing="0"></asp:CheckBoxList>
                                <input id="TPlanIDXHID" type="hidden" value="0" name="TPlanIDXHID" runat="server" size="1">
                            </td>
                        </tr>--%>
                        <%-- <tr id="Check_TR" runat="server">
                            <td class="bluecol">查詢範圍</td>
                            <td class="whitecol">
                                <input id="CheckData" type="checkbox" runat="server">統計全轄區</td>
                        </tr>--%>
                        <%-- <tr id="Org_TR" runat="server">
                            <td class="bluecol">訓練機構</td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <input id="PlanID" type="hidden" name="PlanID" runat="server">
                                <asp:Button ID="Button3" runat="server" Text="查詢班級" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>--%>
                        <%--<tr id="Class_TR" runat="server">
                            <td class="bluecol">班別</td>
                            <td class="whitecol">
                                <asp:ListBox ID="OCID" runat="server" Width="60%" SelectionMode="Multiple" Rows="6"></asp:ListBox>
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>(按Ctrl可以複選班級)
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Visible="False"></asp:TextBox><asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Visible="False"></asp:TextBox>
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                            </td>
                        </tr>--%>
                        <tr>
                            <td class="bluecol">開訓區間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ~
                                <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓區間</td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ~
                                <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <%-- <tr>
                            <td class="bluecol">調查表版本</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblprtType1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Selected="True" Value="A2">原</asp:ListItem>
                                    <asp:ListItem Value="A16">2016年5月</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>--%>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="2" align="center" >
                                <asp:Button ID="Query" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <div align="center" class="whitecol">
        </div>
        <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <div id="Div1" runat="server">
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AllowPaging="True" AutoGenerateColumns="False">
                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center" />
                            <Columns>
                                <asp:BoundColumn HeaderText="序號">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="CTName" HeaderText="縣市別">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="CYCLTYPE" HeaderText="期別">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="STDate" HeaderText="開訓日期">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="TOTAL" HeaderText="結訓人數">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="NUM1" HeaderText="填寫人數">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Q1_Average" HeaderText="一、課程安排 ">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Q2_Average" HeaderText="二、師資、助教及教學">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Q3_Average" HeaderText="三、設備和教材">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Q4_Average" HeaderText="四、行政措施、平均滿意度">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Average" HeaderText="平均滿意度">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="功能">
                                    <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <ItemTemplate>
                                        <asp:Button ID="BtnDetail" runat="server" Text="列印明細" CommandName="Detail" ToolTip="填寫人數為0則無明細" CssClass="asp_Export_M"></asp:Button>
                                    </ItemTemplate>
                                </asp:TemplateColumn>

                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                    </div>
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
                <td align="center" colspan="2" class="whitecol">
                    <asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    <asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button><br>
                    <font class="font" color="red">班級未做結訓作業,無法於匯出、列印資料中顯示該班的滿意度調查統計資料,請確實完成班級結訓作業.</font>
                </td>
            </tr>
        </table>

        <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
        <input id="PlanID" type="hidden" name="PlanID" runat="server" />
    </form>
</body>
</html>
