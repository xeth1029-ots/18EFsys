<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_002.aspx.vb" Inherits="WDAIIP.SD_11_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_11_002</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function GETvalue() {
            document.getElementById('Button7').click();
        }

        function search1() {
            document.form1.hidSearchTag.value = 'search';
            if (document.form1.OCIDValue1.value == '') {
                alert('請選擇職類班別!');
                return false;
            }
        }

        function choose_class() {
            if (document.getElementById('OCID1').values == '') { document.getElementById('Button7').click(); }
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <input id="check_add" type="hidden" name="check_add" runat="server">
        <input id="check_search" type="hidden" name="check_search" runat="server">
        <input id="check_mod" type="hidden" name="check_mod" runat="server">
        <input id="check_del" type="hidden" name="check_del" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <%--<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;就業輔導問卷&gt;&gt;訓練成效追蹤調查表</asp:Label>--%>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;訓練成效與滿意度&gt;&gt;訓練成效追蹤調查表</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%" cellspacing="1" cellpadding="1">
            <tr>
                <td class="bluecol" style="width: 20%">訓練機構</td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    <asp:Button ID="Button7" Style="display: none" runat="server" Text="Button7"></asp:Button>
                    <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">職類/班別</td>
                <td class="whitecol">
                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                    <input id="TMIDValue1" style="width: 40px; height: 22px" type="hidden" name="TMIDValue1" runat="server">
                    <input id="OCIDValue1" style="width: 32px; height: 22px" type="hidden" name="OCIDValue1" runat="server">
                    <input id="hidSearchTag" type="hidden" name="hidSearchTag" runat="server">
                    <span id="HistoryList" style="display: none; left: 28%; position: absolute">
                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
        </table>
        <div class="whitecol" align="center">
            <asp:Button ID="search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
            <asp:Button ID="Print" runat="server" Text="列印空白表" CssClass="asp_Export_M"></asp:Button>
        </div>
        <br>
        <asp:Panel ID="Panel1" runat="server" Width="100%">
            <div style="text-align: center;">
                <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
            </div>
            <div>
                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                    <HeaderStyle CssClass="head_navy" Width="20%"></HeaderStyle>
                    <ItemStyle HorizontalAlign="Center" />
                    <Columns>
                        <asp:BoundColumn HeaderText="序號">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn HeaderText="班別">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="total" HeaderText="結訓人數">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:BoundColumn DataField="num" HeaderText="填寫人數">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        </asp:BoundColumn>
                        <asp:TemplateColumn HeaderText="功能">
                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            <ItemTemplate>
                                <asp:Button ID="Button1" runat="server" Text="查詢" CommandName="view" CssClass="asp_button_M"></asp:Button>
                            </ItemTemplate>
                        </asp:TemplateColumn>
                        <asp:BoundColumn Visible="False" DataField="FTDate" HeaderText="結訓日期"></asp:BoundColumn>
                        <asp:BoundColumn Visible="False" DataField="CyclType" HeaderText="CyclType"></asp:BoundColumn>
                        <asp:BoundColumn Visible="False" DataField="LevelType" HeaderText="LevelType"></asp:BoundColumn>
                    </Columns>
                </asp:DataGrid>
            </div>
            <div>
                <asp:Label ID="Label1" runat="server" CssClass="font"></asp:Label>
            </div>
            <div style="margin-top: 3px; margin-bottom: 3px" align="center">
                <asp:Label ID="msg2" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
            </div>
            <div>
                <table id="StudentTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                    <tr>
                        <td>
                            <asp:DataGrid ID="DG_stud" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                <HeaderStyle CssClass="head_navy" Width="25%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" />
                                <Columns>
                                    <asp:BoundColumn HeaderText="學號">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn DataField="Name" HeaderText="姓名(離退訓日期)">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:BoundColumn HeaderText="填寫狀態">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundColumn>
                                    <asp:TemplateColumn HeaderText="功能">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        <ItemTemplate>
                                            <asp:Button ID="Button4" runat="server" Text="新增" CommandName="insert" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="Button5" runat="server" Text="查詢" CommandName="check" CssClass="asp_button_M"></asp:Button>
                                            <asp:Button ID="Button6" runat="server" Text="清除重填" CommandName="clear" CssClass="asp_button_M"></asp:Button>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
                                    <asp:BoundColumn Visible="False" DataField="StudentID" HeaderText="StudentID"></asp:BoundColumn>
                                </Columns>
                            </asp:DataGrid>
                        </td>
                    </tr>
                </table>
            </div>
        </asp:Panel>
    </form>
</body>
</html>
