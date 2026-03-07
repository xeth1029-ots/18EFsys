<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_011_add.aspx.vb" Inherits="WDAIIP.SD_05_011_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd HTML 4.0 transitional//EN">
<html>
<head>
    <title>SD_05_011_add</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function chkdata() {
            var msg = '';
            //if ($('#YearList option:selected').val() == '0' || $('#YearList option:selected').val() == '') { msg += '請選擇年度!\n'; }
            //if ($('#DistrictList option:selected').val() == '') { msg += '請選擇轄區中心!\n'; }
            //if ($('#Plan_List option:selected').val() == '') { msg += '請選擇訓練計畫!\n'; }

            if (document.form1.YearList.selectedIndex == '0') msg = msg + '請選擇年度!\n';
            //if (document.form1.DistrictList.selectedIndex == '0') msg = msg + '請選擇轄區中心!\n';
            if (document.form1.DistrictList.selectedIndex == '0') msg = msg + '請選擇轄區分署!\n';
            if (document.form1.Plan_List.selectedIndex == '0') msg = msg + '請選擇訓練計畫!\n';
            if (document.form1.ClassName.value == '') msg += '請輸入班級名稱!\n'
            if (document.form1.CosUnit.value == '') msg += '請輸入主辦單位!\n'
            if (document.form1.Name.value == '') msg += '請輸入姓名!\n'
            if (document.form1.SID.value == '') msg = msg + '請輸入身分證號碼!\n';
            else if (!checkId(document.form1.SID.value)) msg += '身分證號碼錯誤\n';
            if (getValue("Sex_List") == '') msg += '請選擇性別!\n'

            if (document.form1.SDate.value != '') {
                if (!checkDate(document.form1.SDate.value)) msg = msg + '開訓日期格式不正確!\n';
            }
            if (document.form1.EDate.value != '') {
                if (!checkDate(document.form1.EDate.value)) msg = msg + '結訓日期格式不正確!\n';
            }
            if (document.form1.birthday.value != '') {
                if (!checkDate(document.form1.birthday.value)) msg = msg + '出生日期格式不正確!\n';
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%--<table class="font" width="600">
        <tr>
            <td class="font">
                首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">90-93歷史資料</font><font color="#990000"><font color="#000000">(<font face="新細明體"><font color="#ff0000">*</font>為必填欄位</font>)<input id="Re_ID" type="hidden" name="Re_ID" runat="server"></font></font>
            </td>
        </tr>
    </table>--%>
        <br>
        <table class="table_nw" id="Table1" cellspacing="1" cellpadding="1" width="100%" runat="server">
            <tr>
                <td class="bluecol_need" style="width: 20%">年度
                </td>
                <td class="whitecol" style="width: 30%">
                    <asp:DropDownList ID="YearList" runat="server">
                        <asp:ListItem Value="0">===請選擇===</asp:ListItem>
                        <asp:ListItem Value="2001">2001</asp:ListItem>
                        <asp:ListItem Value="2002">2002</asp:ListItem>
                        <asp:ListItem Value="2003">2003</asp:ListItem>
                        <asp:ListItem Value="2004">2004</asp:ListItem>
                    </asp:DropDownList>
                </td>
                <%--<td class="bluecol_need" style="width: 20%">轄區中心</td>--%>
                <td class="bluecol_need" style="width: 20%">轄區分署</td>
                <td class="whitecol" style="width: 30%">
                    <asp:DropDownList ID="DistrictList" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td id="td2" runat="server" class="bluecol_need">訓練計畫
                </td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="Plan_List" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">班級名稱
                </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="ClassName" runat="server" Width="60%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">主辦單位
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="CosUnit" Width="95%" runat="server"></asp:TextBox>
                </td>
                <td id="td1" runat="server" class="bluecol">培訓單位
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="trinUnit" Width="95%" runat="server"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓日期
                </td>
                <td class="whitecol" runat="server">
                    <asp:TextBox ID="SDate" runat="server" Width="40%"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                </td>
                <td class="bluecol">結訓日期
                </td>
                <td class="whitecol" runat="server">
                    <asp:TextBox ID="EDate" runat="server" Width="40%"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= EDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">姓名
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="Name" runat="server" Width="40%"></asp:TextBox>
                </td>
                <td class="bluecol_need">身分證字號
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="SID" runat="server" Width="40%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol">生日
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="birthday" runat="server" Width="40%"></asp:TextBox>
                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= birthday.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                </td>
                <td class="bluecol_need">性別
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Sex_List" CssClass="font" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Value="M">男</asp:ListItem>
                        <asp:ListItem Value="F">女</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">身分
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="Ident" runat="server" Width="40%"></asp:TextBox>
                </td>
                <td class="bluecol">連絡電話
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="Tel" runat="server" Width="40%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol">聯絡地址
                </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="Addr" runat="server" Width="95%"></asp:TextBox>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <input id="Button1" type="button" value="回上一頁" name="Button1" runat="server" class="asp_button_M">
                    <%--<asp:ValidationSummary ID="Summary" runat="server" ShowMessageBox="true" ShowSummary="False" DisplayMode="List"></asp:ValidationSummary>--%>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
