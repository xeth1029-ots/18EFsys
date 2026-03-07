<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_002_Wrong.aspx.vb" Inherits="WDAIIP.TC_01_002_Wrong" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>錯誤資料</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%--
        <table id="FrameTable" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
		    <tr>
			    <td>
				    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
				    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000">訓練機構設定</font></asp:Label><font color="#990000">&nbsp;</font>
                </td>
		    </tr>
        </table>
        --%>
        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="font" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="Index" HeaderText="第幾筆錯誤" HeaderStyle-Width="20%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ComIDNO" HeaderText="統一編號" HeaderStyle-Width="20%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="GradeDate" HeaderText="評鑑日期" HeaderStyle-Width="20%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱" HeaderStyle-Width="20%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="Reason" HeaderText="原因" HeaderStyle-Width="20%"></asp:BoundColumn>
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
        </table>
    </form>
</body>
</html>
