<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_008.aspx.vb" Inherits="WDAIIP.SYS_04_008" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訪視警示燈率設定</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;訪視警示燈率設定</asp:Label>
                </td>
            </tr>
        </table>
    <table id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
        <tr>
            <td>
                <%--<table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                    <tr>
                        <td>
                            <font face="新細明體">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;訪視警示燈率設定</font>
                        </td>
                    </tr>
                </table>--%>
                <table id="Table2" class="table_sch" cellpadding="1" cellspacing="1" width="100%">
                    <tr>
                        <td rowspan="3" class="bluecol" style="width:20%">
                            缺失警示
                        </td>
                        <td class="bluecol" style="width:10%">
                            綠燈比率
                        </td>
                        <td class="whitecol" >
                            <asp:TextBox ID="GRate1" runat="server" Columns="5" Width="10%"></asp:TextBox>
                            %
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            黃燈比率
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="YRate1" runat="server" Columns="5" Width="10%"></asp:TextBox>%
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            紅燈比率
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="RRate1" runat="server" Columns="5" Width="10%"></asp:TextBox>%
                        </td>
                    </tr>
                    <tr>
                        <td rowspan="3" class="bluecol">
                            訪視次數不足警示
                        </td>
                        <td class="bluecol">
                            綠燈比率
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="GRate2" runat="server" Columns="5" Width="10%"></asp:TextBox>%
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            黃燈比率
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="YRate2" runat="server" Columns="5" Width="10%"></asp:TextBox>%
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            紅燈比率
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="RRate2" runat="server" Columns="5" Width="10%"></asp:TextBox>%
                        </td>
                    </tr>
                </table>
                <table width="100%">
                    <tr>
                        <td align="center" colspan="3" class="whitecol">                            
                            <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
