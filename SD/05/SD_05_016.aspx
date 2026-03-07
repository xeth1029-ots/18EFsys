<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_016.aspx.vb" Inherits="WDAIIP.SD_05_016" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員補助金歷史查詢</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript">
        function GETvalue() {
            document.getElementById('Button6').click();
        }

        function SetOneOCID() {
            document.getElementById('Button7').click();
        }

        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
        }

        function CheckSearch() {
            if (document.getElementById('OCIDValue1').value == '' && document.getElementById('IDNO').value == '' && document.getElementById('Name').value == '') {
                alert('至少要輸入一項條件');
                return false;
            }
        }

        function choose_class() {
            if (document.getElementById('OCID1').value == '') { document.getElementById('Button7').click(); }
            openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;學員補助金歷史查詢</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="Page1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <table class="table_sch" id="Table2" cellpadding="1" cellspacing="1">
                                    <tr id="OrgTR" runat="server">
                                        <td class="bluecol" style="width: 20%">訓練機構 </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="center" runat="server" Width="55%"></asp:TextBox>
                                            <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                            <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                            <asp:Button ID="Button7" Style="display: none" runat="server"></asp:Button>
                                            <asp:Button ID="Button6" Style="display: none" runat="server" Text="Button6"></asp:Button>
                                            <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                                                <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">職類/班別 </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                            <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                            <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                            <input id="Button4" type="button" value="清除" name="Button4" runat="server" class="asp_button_M">
                                            <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                            <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                            <span id="HistoryList" style="display: none; left: 28%; position: absolute">
                                                <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" style="width: 20%">身分證號碼 </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:TextBox ID="IDNO" runat="server" Width="60%"></asp:TextBox></td>
                                        <td class="bluecol" style="width: 20%">學員姓名 </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:TextBox ID="Name" runat="server" Width="30%"></asp:TextBox></td>
                                    </tr>
                                    <tr id="tr_ddl_INQUIRY_S" runat="server">
                                        <td class="bluecol_need">查詢原因</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                                        </td>
                                    </tr>
                                </table>
                                <table width="100%">
                                    <tr>
                                        <td align="center" class="whitecol">
                                            <asp:Label ID="labPageSize" runat="server" DESIGNTIMEDRAGDROP="30" ForeColor="SlateBlue">顯示列數</asp:Label>
                                            <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                            <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="center" class="whitecol">
                                            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                                    </tr>
                                </table>
                                <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                    <tr>
                                        <td>
                                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" AllowPaging="True" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
                                                        <HeaderStyle Width="30%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                                        <HeaderStyle Width="30%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="Birthday" HeaderText="出生日期" DataFormatString="{0:d}">
                                                        <HeaderStyle Width="30%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="10%" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Button ID="Button3" runat="server" Text="檢視" CommandName="view1" CssClass="asp_button_M"></asp:Button>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
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
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="Page2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center" class="whitecol">
                                <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">姓名 </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="LName" runat="server"></asp:Label></td>
                                        <td class="bluecol" style="width: 20%">身分證號碼 </td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:Label ID="LIDNO" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td colspan="4" class="whitecol">
                                            <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#F5F5F5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                                        <HeaderStyle Width="15%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="參訓課程" HeaderStyle-Width="14%">
                                                        <ItemTemplate>
                                                            <asp:Label ID="labClassCName" runat="server" Text=""></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="開訓日期~結訓日期(異動日期)" HeaderStyle-Width="10%">
                                                        <ItemTemplate>
                                                            <asp:Label ID="labSFMTDATE" runat="server" Text=""></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:BoundColumn DataField="SumOfMoney" HeaderText="申請補助金額">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="BudName" HeaderText="預算別">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn HeaderText="審核狀態" HeaderStyle-Width="10%">
                                                        <ItemTemplate>
                                                            <asp:Label ID="labAppliedStatusM" runat="server" Text=""></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:TemplateColumn HeaderText="撥款狀態" HeaderStyle-Width="10%">
                                                        <ItemTemplate>
                                                            <asp:Label ID="labAppliedStatus" runat="server" Text=""></asp:Label>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                    <asp:BoundColumn DataField="StudStatus" HeaderText="訓練狀態">
                                                        <HeaderStyle Width="10%"></HeaderStyle>
                                                    </asp:BoundColumn>
                                                </Columns>
                                            </asp:DataGrid><asp:Label ID="msg2" runat="server"></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <%--撥款通過總額--%>
                                        <td colspan="4" class="whitecol">(補助總額：
											<asp:Label ID="LabTotal" runat="server"></asp:Label>)-(經費審核通過總額：
											<asp:Label ID="LabSumOfMoney" runat="server"></asp:Label>)=(剩餘可用額度：
											<asp:Label ID="RemainSub" runat="server"></asp:Label>)
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="4" class="whitecol">
                                            <asp:Label ID="LabCostDay" runat="server"></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td colspan="4" class="whitecol"><font color="red">
                                            <asp:Label ID="Lab_TipMsg2" runat="server" Text=""></asp:Label></font></td>
                                    </tr>
                                </table>
                                <asp:Button ID="Button5" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
