<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_007.aspx.vb" Inherits="WDAIIP.SYS_03_007" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>學員資料勾稽查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function search() {
            var msg = '';
            if ((document.form1.SDate.value == '') || (document.form1.EDate.value == '')) msg += '請輸入投保日期!!\n';
            if (document.form1.SDate.value != '' && !checkDate(document.form1.SDate.value)) msg += '起始時間格式不正確\n';
            if (document.form1.EDate.value != '' && !checkDate(document.form1.EDate.value)) msg += '終至時間格式不正確\n';
            if (msg != '') {
                alert(msg);
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;學員資料勾稽查詢</asp:Label>
                </td>
            </tr>
        </table>
        <table id="DataGridtable" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
            <%--<tr><td class="font" colspan="3">首頁&gt;&gt;系統管理&gt;&gt;<font color="#990000">學員資料勾稽查詢</font></td></tr>--%>
        </table>
        <table class="table_nw" id="table3" width="100%" cellspacing="1" cellpadding="1">
            <tr>
                <td class="bluecol_need" style="width: 20%">投保日期:
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="SDate" onfocus="this.blur()" Width="15%" runat="server"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('SDate','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" alt="" />～
                <asp:TextBox ID="EDate" onfocus="this.blur()" Width="15%" runat="server"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('EDate','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" alt="" />
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button3" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
        </table>
        <br />
        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="8" BackColor="White" BorderWidth="1px" BorderStyle="None" BorderColor="#CC9966" Height="416px" OnPageIndexChanged="DataGrid1_SelectedIndexChanged" AllowPaging="true">
            <AlternatingItemStyle BackColor="#F5F5F5" />
            <HeaderStyle CssClass="head_navy" />
            <Columns>
                <asp:BoundColumn DataField="Name" HeaderText="姓名" HeaderStyle-Width="25%"></asp:BoundColumn>
                <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼" HeaderStyle-Width="25%"></asp:BoundColumn>
                <asp:BoundColumn DataField="FM" HeaderText="性別" HeaderStyle-Width="25%"></asp:BoundColumn>
                <asp:BoundColumn DataField="changemode" HeaderText="加退保" HeaderStyle-Width="25%"></asp:BoundColumn>
            </Columns>
            <PagerStyle Visible="False" HorizontalAlign="Center" ForeColor="#330099" BackColor="#FFFFCC"></PagerStyle>
        </asp:DataGrid>
        <asp:HiddenField ID="Hid_guid1" runat="server" />
    </form>
</body>
</html>
