<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_10_006_R.aspx.vb" Inherits="WDAIIP.SD_10_006_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>證書補發</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function ChekSearch() {
            if (document.getElementById("IDNO").value == '') {
                alert("請輸入身分證號!!")
                return false;
            }
        }

        function ShowTR() {

            if (document.form1.Type[0].checked == true) {
                document.getElementById('NO1_TR').style.display = '';
                document.getElementById('NO2_TR').style.display = 'none';
                document.getElementById('NO3_TR').style.display = 'none';
                document.getElementById('NO3_TR2').style.display = 'none';
                document.getElementById('MSG_TR').style.display = 'none';
                document.getElementById('DG_TR').style.display = 'none';
            }
            else if (document.form1.Type[1].checked == true) {
                document.getElementById('NO1_TR').style.display = 'none';
                document.getElementById('NO2_TR').style.display = '';
                document.getElementById('NO3_TR').style.display = 'none';
                document.getElementById('NO3_TR2').style.display = 'none';
                document.getElementById('MSG_TR').style.display = 'none';
                document.getElementById('DG_TR').style.display = 'none';
            }
            else if (document.form1.Type[2].checked == true) {
                document.getElementById('NO1_TR').style.display = 'none';
                document.getElementById('NO2_TR').style.display = 'none';
                document.getElementById('NO3_TR').style.display = '';
                document.getElementById('NO3_TR2').style.display = '';
                document.getElementById('MSG_TR').style.display = 'none';
                document.getElementById('DG_TR').style.display = 'none';
            }

        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table1" width="100%">
            <tr>
                <td class="font">
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server"> 首頁&gt;&gt;學員動態管理&gt;&gt;證書及證明管理&gt;&gt;<font color="#990000">證書補發</font> </asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" id="Table2" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol_need" width="18%">身分證號
                </td>
                <td class="whitecol" width="82%">
                    <asp:TextBox ID="IDNO" runat="server" MaxLength="15"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">列印格式
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblYearType1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="1" Selected="True">西元年</asp:ListItem>
                        <asp:ListItem Value="2">民國年</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">證書類型
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Type" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Value="1">在訓證明</asp:ListItem>
                        <asp:ListItem Value="2">受訓證明</asp:ListItem>
                        <asp:ListItem Value="3">結訓證書</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr id="NO1_TR" runat="server">
                <td class="bluecol">在訓證明字號
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="NO1" runat="server" MaxLength="300" Width="60%"></asp:TextBox>
                </td>
            </tr>
            <tr id="NO2_TR" runat="server">
                <td class="bluecol">受訓證明字號
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="NO2" runat="server" MaxLength="300" Width="60%"></asp:TextBox>
                </td>
            </tr>
            <tr id="NO3_TR" runat="server">
                <td class="bluecol">結訓證書字號
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="NO3" runat="server" MaxLength="300" Width="60%"></asp:TextBox>
                </td>
            </tr>
            <tr id="NO3_TR2" runat="server">
                <td class="bluecol">列印格式
                </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="Type2" runat="server" RepeatDirection="Horizontal">
                        <asp:ListItem Value="1" Selected="True">自辦</asp:ListItem>
                        <asp:ListItem Value="2">委辦</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td align="center">
                    <asp:Button ID="Search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr id="MSG_TR" runat="server">
                <td align="center">
                    <asp:Label ID="labMsg" Style="color: red" runat="server">查無資料</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" width="100%" align="center">
            <tr id="DG_TR" runat="server">
                <td colspan="2">
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" AllowPaging="True">
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy" HorizontalAlign="Center" VerticalAlign="Middle"></HeaderStyle>
                        <Columns>
                            <asp:TemplateColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="3%"></asp:TemplateColumn>
                            <asp:BoundColumn DataField="SName" HeaderText="姓名" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="7%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="8%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="birthday" HeaderText="出生日期" DataFormatString="{0:d}" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="7%"></asp:BoundColumn>
                            <%--<asp:BoundColumn DataField="Distid" HeaderText="轄區中心" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="13%"></asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="Distid" HeaderText="轄區分署" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="13%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="years" HeaderText="年度" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="4%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="orgName" HeaderText="訓練機構" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="17%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班別名稱" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="19%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="STUDSTATUS_N" HeaderText="訓練狀態" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="7%"></asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center">
                                <ItemTemplate>
                                    <asp:Button ID="print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                </ItemTemplate>
                                <FooterStyle HorizontalAlign="Center" VerticalAlign="Middle"></FooterStyle>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
