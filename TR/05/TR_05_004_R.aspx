<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_004_R.aspx.vb" Inherits="WDAIIP.TR_05_004_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職業訓練自辦/委辦人數分佈</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript">
        function search() {
            if (document.form1.Syear.selectedIndex == 0) {
                alert('請選擇年度');
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;訓練與就業需求管理&gt;&gt;統計分析&gt;&gt;<FONT color="#990000">職業訓練自辦/委辦人數分佈</FONT>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="Table2" class="table_sch" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol_need" width="100">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Syear" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
