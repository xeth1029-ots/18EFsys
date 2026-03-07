<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_02_001_R.aspx.vb" Inherits="WDAIIP.CP_02_001_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>結訓學員概況統計表</title>
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
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;公務統計報表&gt;&gt;結訓學員概況統計表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td width="20%" class="bluecol_need">結訓日期</td>
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
                <TD  bgColor="#cc6666"><font color="#ffffff">&nbsp;&nbsp;&nbsp; 訓練計畫</font></TD>
				<TD bgColor="#ffecec"><asp:CheckBoxList id="TPlan" CssClass="font" runat="server" RepeatColumns="3"></asp:CheckBoxList></TD>
			</TR>
            --%>
            <tr id="TPlanID0_TR" runat="server">
                <td class="bluecol">訓練計畫(職前)</td>
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
            <tr>
                <td class="bluecol">X軸(複選)</td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="SortX" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="2">
                        <asp:ListItem Value="性別">性別</asp:ListItem>
                        <asp:ListItem Value="年齡">年齡</asp:ListItem>
                        <asp:ListItem Value="教育程度">教育程度</asp:ListItem>
                        <asp:ListItem Value="兵役">兵役</asp:ListItem>
                        <asp:ListItem Value="訓練性質">訓練性質（封面）</asp:ListItem>
                        <asp:ListItem Value="學員身分">學員身分</asp:ListItem>
                        <asp:ListItem Value="參訓前一個月工作情形">參訓前一個月工作情形</asp:ListItem>
                        <asp:ListItem Value="參訓前一個月有否尋找工作">參訓前一個月有否尋找工作</asp:ListItem>
                        <asp:ListItem Value="您參加本次訓練後是否找到工作">您參加本次訓練後是否找到工作</asp:ListItem>
                        <asp:ListItem Value="訓練機構">訓練機構(公訓：機構別／補助地方政府：縣市別)</asp:ListItem>
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">Y軸(複選)</td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="SortY" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="3">
                        <asp:ListItem Value="訓練機構">訓練機構(公訓：機構別／補助地方政府：縣市別)</asp:ListItem>
                        <asp:ListItem Value="訓練性質">訓練性質(封面)</asp:ListItem>
                        <asp:ListItem Value="委託訓練">委託訓練(封面)</asp:ListItem>
                        <asp:ListItem Value="訓練職類">訓練職類(封面)</asp:ListItem>
                        <asp:ListItem Value="上課時段">上課時段(封面)</asp:ListItem>
                        <asp:ListItem Value="學員身分">學員身分</asp:ListItem>
                        <asp:ListItem Value="參訓動機">參訓動機</asp:ListItem>
                        <asp:ListItem Value="結訓後動向">結訓後動向</asp:ListItem>
                        <asp:ListItem Value="參訓前一個月工作情形">參訓前一個月工作情形</asp:ListItem>
                        <asp:ListItem Value="參訓前一個月有否尋找工作">參訓前一個月有否尋找工作</asp:ListItem>
                        <asp:ListItem Value="訓練後或訓練期間有否找到工作">訓練後或訓練期間有否找到工作</asp:ListItem>
                        <asp:ListItem Value="本次訓練後覺得不滿意需要改進的為何">本次訓練後覺得不滿意需要改進的為何</asp:ListItem>
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td colspan="2" align="center"><asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></td>
            </tr>
        </table>
    </form>
</body>
</html>