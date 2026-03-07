<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_001_import.aspx.vb" Inherits="WDAIIP.TC_01_001_import" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>計畫匯入</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function chkdata() {
            var year1 = document.getElementById('Fromyear');
            var year2 = document.getElementById('Toyear');
            var msg = '';
            if (document.getElementById('Table3').style.display == '' && document.form1.DistID.selectedIndex == 0) msg += '請選擇轄區\n';
            if (year1.selectedIndex == 0) msg += '請選擇來源年度\n';
            if (year2.selectedIndex == 0) msg += '請選擇目的年度\n';
            if (year1.value == year2.value) msg += '不能匯入相同的年度!\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
    <%-- <style type="text/css"> #Table3 { margin-bottom: 0px; } </style> --%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="Table3" runat="server" cellpadding="1" cellspacing="1" class="table_nw" width="100%">
            <tr>
                <td class="bluecol_need" width="20%">轄區 </td>
                <td class="whitecol" width="80%">
                    <asp:DropDownList ID="DistID" runat="server" cellpadding="1" cellspacing="1"></asp:DropDownList></td>
            </tr>
        </table>
        <asp:Panel ID="Table1" runat="server">
            <table cellpadding="1" cellspacing="1" class="table_nw" width="100%">
                <tr>
                    <td class="bluecol_need" width="20%">來源年度 </td>
                    <td class="whitecol" width="80%">
                        <asp:DropDownList ID="Fromyear" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="20%">目的年度 </td>
                    <td class="whitecol" width="80%">
                        <asp:DropDownList ID="Toyear" runat="server"></asp:DropDownList></td>
                </tr>
            </table>
            <table width="100%">
                <tr>
                    <td class="whitecol" align="center">
                        <asp:Button ID="Button1" runat="server" Text="匯入" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BTN_BACK1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center" class="whitecol">
                    <input type="button" value="關閉" id="Button2" name="Button2" runat="server" class="asp_button_M"></td>
            </tr>
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" AllowPaging="false" CssClass="font" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="Years" HeaderText="年度">
                                <ItemStyle HorizontalAlign="Center" Width="8%" />
                            </asp:BoundColumn>
                            <%--<asp:BoundColumn DataField="DistID" HeaderText="轄區中心代碼">--%>
                            <asp:BoundColumn DataField="DistID" HeaderText="轄區分署代碼">
                                <ItemStyle HorizontalAlign="Center" Width="12%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TPlanID" HeaderText="訓練計畫代碼">
                                <ItemStyle HorizontalAlign="Center" Width="12%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Seq" HeaderText="序號">
                                <ItemStyle HorizontalAlign="Center" Width="10%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Sponsor" HeaderText="主辦單位 ">
                                <ItemStyle HorizontalAlign="Center" Width="14%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Cosponsor" HeaderText="協辦單位">
                                <ItemStyle HorizontalAlign="Center" Width="14%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="SDate" HeaderText="時效起日" DataFormatString="{0:d}">
                                <ItemStyle HorizontalAlign="Center" Width="10%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="EDate" HeaderText="時效迄日" DataFormatString="{0:d}">
                                <ItemStyle HorizontalAlign="Center" Width="10%" />
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlanKind" HeaderText="計畫種類">
                                <ItemStyle HorizontalAlign="Center" Width="10%" />
                            </asp:BoundColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
