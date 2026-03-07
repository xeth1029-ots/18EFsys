<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_02_014_R.aspx.vb" Inherits="WDAIIP.CP_02_014_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>公共職訓機構結訓人數按特定對象別分</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function search() {
            var msg = '';
            if (isEmpty(document.form1.start_date) && isEmpty(document.form1.end_date)) {
                msg += '請選擇日期範圍!\n';
            }
            if (isEmpty(document.form1.OCID)) {
                msg += '請選擇統計對象!\n';
            }
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
        <table id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;公務統計報表&gt;&gt;<font color="#990000">公立職訓機構特定對象結訓人數</font></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="Table3" class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol_need">結訓日期</td>
                            <td class="whitecol">
                                <span runat="server">
                                    <asp:TextBox ID="start_date" runat="server" Width="100px"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"> ～
                                    <asp:TextBox ID="end_date" runat="server" Width="100px"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">統計對象</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="OCID" runat="server">
                                    <asp:ListItem Value="">===請選擇===</asp:ListItem>
                                    <asp:ListItem Value="0">全部</asp:ListItem>
                                    <%--<asp:ListItem Value="1">局屬全部</asp:ListItem>--%>
                                    <asp:ListItem Value="1">署屬全部</asp:ListItem>
                                    <%--<asp:ListItem Value="2">非局屬全部</asp:ListItem>--%>
                                    <asp:ListItem Value="2">非署屬全部</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <%--
					    <TR>
							<TD width="100" bgColor="#cc6666"><font color="#ffffff">&nbsp;&nbsp;&nbsp; 訓練計畫</font></TD>
							<TD bgColor="#ffecec"><asp:CheckBoxList id="TPlan" CssClass="font" runat="server" RepeatColumns="3"></asp:CheckBoxList></TD>
					    </TR>
                        --%>
                        <tr id="TPlanID0_TR" runat="server">
                            <td width="100" class="bluecol">訓練計畫(職前)</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chkTPlanID0" runat="server" CellSpacing="0" CellPadding="0" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font"></asp:CheckBoxList>
                                <input id="TPlanID0HID" value="0" type="hidden" name="TPlanID0HID" runat="server">
                            </td>
                        </tr>
                        <tr id="TPlanID1_TR" runat="server">
                            <td class="bluecol">訓練計畫(在職)</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chkTPlanID1" runat="server" CellSpacing="0" CellPadding="0" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font"></asp:CheckBoxList>
                                <input id="TPlanID1HID" value="0" type="hidden" name="TPlanID1HID" runat="server">
                            </td>
                        </tr>
                        <tr id="TPlanIDX_TR" runat="server">
                            <td class="bluecol">訓練計畫(其他)</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="chkTPlanIDX" runat="server" CellSpacing="0" CellPadding="0" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font"></asp:CheckBoxList>
                                <input id="TPlanIDXHID" value="0" type="hidden" name="TPlanIDXHID" runat="server">
                            </td>
                        </tr>
                    </table>
                    <div align="center"><asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>