<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_13_003_Bligate.aspx.vb" Inherits="WDAIIP.SD_13_003_Bligate" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>補助審核</title>
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
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;補助金請領&gt;&gt;補助審核</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table6" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td width="65">姓&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;名： </td>
                            <td width="120">
                                <asp:Label ID="Name" runat="server"></asp:Label>
                            </td>
                            <td align="right" width="80">身分證號碼： </td>
                            <td>
                                <asp:Label ID="IDNO" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid3" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="font">
                        <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn HeaderText="序號"></asp:BoundColumn>
                            <asp:BoundColumn DataField="FType" HeaderText="保險別"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ActNo" HeaderText="保險證號"></asp:BoundColumn>
                            <asp:BoundColumn DataField="MDate" HeaderText="異動日期" DataFormatString="{0:d}"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ChangeMode" HeaderText="異動狀況"></asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="Salary" HeaderText="投保薪資"></asp:BoundColumn>
                            <asp:BoundColumn DataField="UName" HeaderText="投保單位"></asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Label ID="msg3" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Button ID="Button5" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
        </table>

        <asp:HiddenField ID="HidActNo" runat="server" />
        <asp:HiddenField ID="HidSTDate" runat="server" />
        <asp:HiddenField ID="HidSOCID" runat="server" />
    </form>
</body>
</html>
