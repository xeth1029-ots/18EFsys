<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_002.aspx.vb" Inherits="WDAIIP.CM_03_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>CM_03_002</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript">
        function search() {
            if (document.form1.Syear.selectedIndex == 0) {
                alert('請先選擇年度');
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練經費控管&gt;&gt;經費查詢&gt;&gt;結訓人數及補助金額_依轄區</asp:Label>
                </td>
            </tr>
        </table>


        <table class="font" id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>

                    <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="20%">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Syear" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練計畫
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="TPlanID" TabIndex="3" runat="server" RepeatDirection="Horizontal"
                                    CssClass="font" RepeatColumns="3">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Table ID="ShowDataTable" runat="server" Width="100%" CssClass="font">
                    </asp:Table>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
