 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_04_019.aspx.vb" Inherits="WDAIIP.TR_04_019" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_04_019</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript">
        function chkSearch() {
            var msg = '';
            if (document.form1.Syear.selectedIndex == 0) msg += '請選擇年度\n';
            //if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區中心\n';
            if (document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區分署\n';
            if (document.form1.FTDate1.value != '') {
                if (!checkDate(document.form1.FTDate1.value)) msg += '結訓期間 的起始日不是正確的日期格式\n';
            }
            if (document.form1.FTDate2.value != '') {
                if (!checkDate(document.form1.FTDate2.value)) msg += '結訓期間 的迄止日不是正確的日期格式\n';
            }
            var obj = '';
            var num = 0;
            var j = 0;
            obj = 'TPlanID';
            num = getCheckBoxListValue(obj).length
            j = 0;
            for (var i = 1; i < num; i++) {
                var mycheck = document.getElementById(obj + '_' + i);
                if (mycheck.checked) { j += 1; }
            }
            if (j == 0) msg += '請選擇 訓練計畫\n';

            obj = 'IsGetJob';
            num = getCheckBoxListValue(obj).length
            j = 0;
            for (var i = 0; i < num; i++) {
                var mycheck = document.getElementById(obj + '_' + i);
                if (mycheck.checked) { j += 1; }
            }
            if (j == 0) msg += '請選擇 輔導就業現況\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

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
    <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr> 
            <td>
                <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                            <asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;訓練與就業需求管理&gt;&gt;就業追蹤報表&gt;&gt;<FONT color="#990000">就服輔導就業名單</FONT>
                            </asp:Label>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                    <tr>
                        <td class="bluecol_need">
                            年度
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="Syear" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <%--<td class="bluecol_need" width="100">轄區中心</td>--%>
                        <td class="bluecol_need" width="100">轄區分署</td>
                        <td class="whitecol">
                            <asp:DropDownList ID="DistID" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            結訓期間
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar2('FTDate1','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            <font color="#000000">～</font>
                            <asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox>
                            <img style="cursor: pointer" onclick="javascript:show_calendar2('FTDate2','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            訓練計畫
                        </td>
                        <td class="whitecol">
                            <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="0" CellPadding="0" RepeatColumns="3">
                            </asp:CheckBoxList>
                            <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server" size="1">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need">
                            輔導就業現況
                        </td>
                        <td class="whitecol">
                            <asp:CheckBoxList ID="IsGetJob" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="0" CellPadding="0" RepeatColumns="3">
                                <asp:ListItem Value="1">已就業</asp:ListItem>
                                <asp:ListItem Value="0">未就業</asp:ListItem>
                                <asp:ListItem Value="2">不就業</asp:ListItem>
                            </asp:CheckBoxList>
                        </td>
                    </tr>
                </table>
                <p align="center">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                    <asp:TextBox ID="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:TextBox>
                    <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
                    <asp:Button ID="btnExport1" runat="server" Text="匯出Excel" CssClass="asp_Export_M"></asp:Button>
                </p>
                <p align="center">
                    <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
                </p>
            </td>
        </tr>
    </table>
    <table id="ResultTable" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td>
                <div id="Div1" runat="server">
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AllowPaging="True">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </div>
            </td>
        </tr>
        <tr>
            <td align="center">
                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
            </td>
        </tr>
    </table>
    </form>
    <%--
		<TR>
			<TD>
			</TD>
		</TR>
		<TR>
			<TD>
			</TD>
		</TR>
    --%>
    <asp:Label Style="z-index: 0" ID="Label1" runat="server" CssClass="font" ForeColor="Red">範圍：排除離退訓學員</asp:Label>
</body>
</html>
