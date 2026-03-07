<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_08_002.aspx.vb" Inherits="WDAIIP.SYS_08_002" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SYS_08_002</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function Check() {
            var msg = '';
            if (document.getElementById('Name3').value == '') {
                msg = '請輸入【分類標題】!\n';
            }
            if (document.getElementById('Serial').value == '') {
                msg += '請輸入【排序序號】!\n';
            }
            if (!isUnsignedInt(document.getElementById('Serial').value)) {
                msg += '【排序序號】必須為數字!\n';
            }
            if (msg != "") { alert(msg); }
        }



    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>首頁&gt;&gt;系統管理&gt;&gt;問卷管理&gt;&gt;問卷分類標題設定
                </td>
            </tr>
        </table>
        <table class="font" id="table_F" width="100%" runat="server">
            <tr>
                <td class="SD_TD1" style="width: 82px" align="center">
                    <label>
                        <span class="style3">問卷名稱</span></label>
                </td>
                <td class="SD_TD2">
                    <input id="Ipt_Name" style="width: 373px; height: 22px" type="text" maxlength="100" size="56" name="Ipt_Name" runat="server" height="18">
                </td>
            </tr>
            <tr align="center">
                <td colspan="2">
                    <input id="search" type="button" value="查詢" name="search" runat="server">
                </td>
            </tr>
        </table>
        <table id="Table2" width="100%" runat="server">
            <tr width="100%">
                <td align="center" width="100%">
                    <font face="新細明體">
                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" runat="server">
                            <AlternatingItemStyle HorizontalAlign="Center" BackColor="White"></AlternatingItemStyle>
                            <ItemStyle HorizontalAlign="Center" BackColor="#EBF8FF"></ItemStyle>
                            <HeaderStyle HorizontalAlign="Center" ForeColor="White" BackColor="#2AAFC0"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn HeaderText="序號">
                                    <HeaderStyle Width="10%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Name" HeaderText="問卷名稱">
                                    <HeaderStyle Width="65%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Avail" HeaderText="狀態">
                                    <HeaderStyle Width="15%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="功能">
                                    <HeaderStyle Width="10%"></HeaderStyle>
                                    <ItemTemplate>
                                        <asp:Button ID="Btn_edit" runat="server" Text="設定問卷分類標題" CommandName="edit1"></asp:Button>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn Visible="False" DataField="SVID" HeaderText="SVID"></asp:BoundColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid></font><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
            <tr width="100%">
                <td width="100%">
                    <p align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </p>
                </td>
            </tr>
        </table>
        <table class="font" id="Table3" width="100%" border="0" runat="server">
            <tr>
                <td class="SD_TD1" style="width: 99px" align="center">
                    <label>
                        問卷名稱</label>
                </td>
                <td class="SD_TD2">
                    <label id="QName" runat="server">
                    </label>
                </td>
            </tr>
            <tr>
                <td class="SD_TD1" style="width: 99px" align="center">
                    <label>
                        分類標題</label>
                </td>
                <td class="SD_TD2">
                    <input id="Name3" style="width: 392px; height: 22px" size="60" runat="server">
                </td>
            </tr>
            <tr>
                <td class="SD_TD1" style="width: 99px" align="center">
                    <label>
                        排序序號</label>
                </td>
                <td class="SD_TD2">
                    <input id="Serial" style="width: 32px; height: 22px" runat="server">
                    <input id="SVID2" style="width: 32px; height: 22px" type="hidden" runat="server">
                    <input id="MODE" style="width: 32px; height: 22px" type="hidden" name="MODE" runat="server"><input id="SKID2" style="width: 32px; height: 22px" type="hidden" name="SKID2" runat="server">
                </td>
            </tr>
            <tr align="center">
                <td colspan="2">
                    <input id="Save" type="button" value="儲存" name="Save" runat="server"><input id="Return1" type="button" value="回上一頁" name="Return1" runat="server">
                </td>
            </tr>
        </table>
        <table id="Table4" width="100%" runat="server">
            <tr width="100%">
                <td align="center" width="100%">
                    <font face="新細明體">
                        <asp:DataGrid ID="Datagrid2" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" runat="server">
                            <AlternatingItemStyle HorizontalAlign="Center" BackColor="White"></AlternatingItemStyle>
                            <ItemStyle HorizontalAlign="Center" BackColor="#EBF8FF"></ItemStyle>
                            <HeaderStyle HorizontalAlign="Center" ForeColor="White" BackColor="#2AAFC0"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn DataField="Serial" HeaderText="序號">
                                    <HeaderStyle Width="5%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="Topic" HeaderText="分類標題">
                                    <HeaderStyle Width="75%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Left" VerticalAlign="Middle"></ItemStyle>
                                </asp:BoundColumn>
                                <asp:TemplateColumn HeaderText="功能">
                                    <HeaderStyle Width="20%"></HeaderStyle>
                                    <ItemTemplate>
                                        <asp:Button ID="edit" runat="server" Text="修改" CommandName="edit"></asp:Button>
                                        <asp:Button ID="del" runat="server" Text="刪除" CommandName="del"></asp:Button>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn Visible="False" DataField="SKID" HeaderText="SKID"></asp:BoundColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid></font><uc1:PageControler ID="PageControler2" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <font face="新細明體"></font>
    </form>
</body>
</html>
