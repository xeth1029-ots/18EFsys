<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_006_R.aspx.vb" Inherits="WDAIIP.SD_01_006_R" %>

 

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>e網報名審核</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;e網審核未審明細表</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="datagridtable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowSorting="true" CssClass="font" AutoGenerateColumns="false" AllowPaging="True" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn DataField="orgname" HeaderText="訓練單位">
                                <ItemStyle HorizontalAlign="Center" Width="30%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="classcname" HeaderText="班別名稱">
                                <ItemStyle HorizontalAlign="Center" Width="40%"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="needcheck" HeaderText="未審人數">
                                <ItemStyle HorizontalAlign="Center" Width="30%"></ItemStyle>
                            </asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <br />
        <div align="center" class="whitecol">
            <asp:Button ID="print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
        </div>
        <center>&nbsp;<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></center>
    </form>
</body>
</html>