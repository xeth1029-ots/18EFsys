<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_001.aspx.vb" Inherits="WDAIIP.SYS_06_001" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>刪除記錄查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
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
            if (!isChecked(document.getElementsByName('FunID'))) msg += '請選擇功能項目\n';
            if (!CheckDateValue(document.getElementById('SDate').value)) msg += '起始日期時間格式不正確\n';
            if (!CheckDateValue(document.getElementById('EDate').value)) msg += '結束日期時間格式不正確\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function CheckDateValue(MyDate) {
            if (MyDate != '') {
                if (!checkDate(MyDate)) return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">

        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;刪除記錄查詢</asp:Label>
                </td>
            </tr>
        </table>

        <table id="Frametable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">

                    <table class="table_nw" id="table2" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">日期區間
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="SDate" runat="server" Columns="8" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('SDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
                                <asp:TextBox ID="EDate" runat="server" Columns="8" Width="15%"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('EDate','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">功能項目
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="FunID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                    <asp:ListItem Value="57">訓練機構設定</asp:ListItem>
                                    <asp:ListItem Value="60">開班資料轉入</asp:ListItem>
                                    <asp:ListItem Value="63">計畫查詢</asp:ListItem>
                                    <asp:ListItem Value="83">學員資料維護</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">帳號
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ModifyAcct" runat="server" Width="30%"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridtable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" AllowPaging="true" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="ModifyAcct" HeaderText="帳號">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ACCNAME" HeaderText="姓名">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ModifyDate" HeaderText="紀錄時間">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="DelNote" HeaderText="說明">
                                            <HeaderStyle Width="70%"></HeaderStyle>
                                        </asp:BoundColumn>
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
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>

    </form>
</body>
</html>
