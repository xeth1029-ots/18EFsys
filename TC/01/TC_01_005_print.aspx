<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="TC_01_005_print.aspx.vb" Inherits="WDAIIP.TC_01_005_print" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TC_01_005_print</title>
    <meta name="vs_showGrid" content="False">
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    &nbsp;
    <table class="font" cellspacing="1" cellpadding="1" border="0" width="100%">
        <%--<tr>
            <td>
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                            <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;課程代碼列印</asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>--%>
        <tr>
            <td>
                <table class="table_sch" cellspacing="1" cellpadding="1" border="0" width="100%">
                    <tr>
                        <td id="Td3" runat="server" class="bluecol" style="width:20%">
                            年度
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="yearlist" runat="server" >
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td id="Td2" runat="server" class="bluecol">
                            訓練職類
                        </td>
                        <td class="whitecol">                            
                            <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                            <input id="career" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="career" runat="server">
                            <input id="trainValue" type="hidden" name="trainValue" runat="server" size="1">                           
                        </td>
                    </tr>
                    <tr>
                        <td id="Td1" runat="server" class="bluecol">
                            課程關鍵字
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="CourseName" runat="server" MaxLength="50" Width="40%"></asp:TextBox>
                        </td>
                    </tr>
                    <tr>
                        <td id="Td4" runat="server" class="bluecol">
                            列印方式
                        </td>
                        <td >                            
                            <asp:RadioButtonList ID="RadioButtonList1" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                <asp:ListItem Value="0" Selected="True">排課匯入</asp:ListItem>
                                <asp:ListItem Value="1">課程編碼</asp:ListItem>
                            </asp:RadioButtonList>                            
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <div align="center" class="whitecol">                                
                                <asp:Button ID="print" runat="server" Text="列印" class="asp_Export_M"></asp:Button>
                                <asp:Button ID="Button1" runat="server" Text="回上一頁" class="asp_button_M"></asp:Button>                                
                            </div>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
