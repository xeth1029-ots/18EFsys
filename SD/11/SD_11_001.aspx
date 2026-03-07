<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_11_001.aspx.vb" Inherits="WDAIIP.SD_11_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練期末學員滿意度</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() { document.getElementById('Button7').click(); }

        function search1() {
            //document.form1.hidSearchTag.value = 'search';
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
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;訓練成效與滿意度&gt;&gt;訓練期末學員滿意度</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol_need" style="width: 20%">訓練機構</td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" Width="55%" onfocus="this.blur()"></asp:TextBox>
                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini">
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    <asp:Button ID="Button7" Style="display: none" runat="server" Text="Button7"></asp:Button>
                    <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">職類/班別</td>
                <td class="whitecol">
                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                    <input id="TMIDValue1" style="width: 40px; height: 22px" type="hidden" name="TMIDValue1" runat="server">
                    <input id="OCIDValue1" style="width: 32px; height: 22px" type="hidden" name="OCIDValue1" runat="server">
                    <span id="HistoryList" style="position: absolute; display: none; left: 30%;">
                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr id="trButton13" runat="server">
                <td class="bluecol">匯入問卷資料</td>
                <td class="whitecol">
                    <input id="File1" type="file" name="File1" runat="server" size="44" accept=".csv,.ods" />
                    <asp:Button ID="Button13" runat="server" Text="匯入資料卡" CssClass="asp_button_M"></asp:Button>(必須為csv或ods 格式)<br />
                    <asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" NavigateUrl="../../Doc/StudQues_v21.zip" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                </td>
            </tr>
            <tr id="trTrnPre1" runat="server">
                <td class="bluecol">職前調查表版本</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblprtType1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Selected="True" Value="A2">原</asp:ListItem>
                        <asp:ListItem Value="A16">2016年5月</asp:ListItem>
                    </asp:RadioButtonList>&nbsp;&nbsp;&nbsp;
                </td>
            </tr>
            <tr>
                <td class="whitecol" colspan="2" align="center">
                    <asp:Button ID="search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="Print" runat="server" Text="列印空白表" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
        </table>
        <br>
        <asp:Panel ID="Panel1" runat="server" Width="100%">
            <div align="center">
                <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy" Width="20%"></HeaderStyle>
                <ItemStyle HorizontalAlign="Center" />
                <Columns>
                    <asp:BoundColumn HeaderText="序號">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ClassCName2" HeaderText="班別">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="total" HeaderText="結訓人數">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="num1" HeaderText="填寫人數">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn HeaderText="功能">
                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        <ItemTemplate>
                            <asp:Button ID="btnView1" runat="server" Text="查詢" CssClass="asp_button_M" CommandName="view"></asp:Button>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
            </asp:DataGrid>
            <div>
                <asp:Label ID="Label1" runat="server" CssClass="font"></asp:Label>
            </div>
            <div style="margin-top: 3px; margin-bottom: 3px" align="center">
                <asp:Label ID="msg2" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
            <table id="StudentTable" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                <tr>
                    <td>
                        <asp:DataGrid ID="DG_stud" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy" Width="25%"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Center" />
                            <Columns>
                                <asp:BoundColumn DataField="StudentID2" HeaderText="學號">
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
                                        <asp:Button ID="BtnAdd4" runat="server" Text="新增" CommandName="insert" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="BtnEdit" runat="server" Text="修改" CommandName="Edit" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="BtnQry5" runat="server" Text="查詢" CommandName="check" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="BtnClear6" runat="server" Text="清除重填" CommandName="clear" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="BtnPrint" runat="server" Text="列印" CommandName="print" CssClass="asp_Export_M"></asp:Button>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                        </asp:DataGrid>
                    </td>
                </tr>
                <tr>
                    <td align="center"></td>
                </tr>
            </table>
        </asp:Panel>
        <%--<input id="hidSearchTag" type="hidden" name="hidSearchTag" runat="server">--%>
        <%--<input id="check_search" type="hidden" size="5" name="check_search" runat="server">
	    <input id="check_add" type="hidden" size="5" name="check_add" runat="server">
	    <input id="check_mod" type="hidden" size="5" name="check_mod" runat="server">
	    <input id="check_del" type="hidden" size="5" name="check_del" runat="server">--%>
        <asp:HiddenField ID="Hid_rblprtType1" runat="server" />
    </form>
</body>
</html>
