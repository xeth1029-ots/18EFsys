<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_004.aspx.vb" Inherits="WDAIIP.SYS_06_004" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>在職補助金設定</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <%--    <link rel="stylesheet" type="text/css" href="../../style.css">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/common.js"></script>--%>
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
        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }
    </script>
    <style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 32px; }
        .auto-style2 { color: #333333; padding: 4px; height: 32px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;在職補助金設定</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>

                    <asp:Panel ID="panelSearch" runat="server">
                        <table id="Table3" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">&nbsp; 啟用日期
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="STDate_s" onfocus="this.blur()" runat="server" MaxLength="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate_s.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">&nbsp;&nbsp;啟用狀態
                                </td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rblisUsed_s" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="Y" Selected="True">是</asp:ListItem>
                                        <asp:ListItem Value="N">否</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button Style="z-index: 0" ID="btnAddnew" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table id="DataGridTable1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td align="center">
                                    <asp:DataGrid Style="z-index: 0" ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CssClass="font" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號">
                                                <HeaderStyle Width="5%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDate" HeaderText="啟用日期">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="計算基準">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lbPYMs" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="適用計畫">
                                                <HeaderStyle Width="55%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lbPlanNames" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle Width="20%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="btnView1" runat="server" Text="檢視" CommandName="View1" CssClass="linkbutton"></asp:LinkButton>
                                                    <asp:LinkButton ID="btnEdit1" runat="server" Text="修改" CommandName="Edit1" CssClass="linkbutton"></asp:LinkButton>
                                                    <asp:LinkButton ID="btnDelete1" runat="server" Text="刪除" CommandName="Delete1" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 31px" align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="PanelEdit1" runat="server">
                        <table id="Table4" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">啟用日期<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="STDate" onfocus="this.blur()" runat="server" MaxLength="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">計算基準<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox Style="z-index: 0" ID="PYears" runat="server" MaxLength="2" Columns="2" Width="8%"></asp:TextBox>年
                                <asp:TextBox Style="z-index: 0" ID="PMoneys" runat="server" MaxLength="2" Columns="2" Width="10%"></asp:TextBox>萬
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">啟用狀態<font color="red">*</font>
                                </td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rblisUsed" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="Y" Selected="True">是</asp:ListItem>
                                        <asp:ListItem Value="N">否</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr id="TPlanID0_TR" runat="server">
                                <td class="auto-style1">訓練計畫(職前)
                                </td>
                                <td class="auto-style2">
                                    <asp:CheckBoxList ID="chkTPlanID0" runat="server" CellSpacing="0" CellPadding="0" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font">
                                    </asp:CheckBoxList>
                                    <input id="TPlanID0HID" value="0" type="hidden" name="TPlanID0HID" runat="server">
                                </td>
                            </tr>
                            <tr id="TPlanID1_TR" runat="server">
                                <td class="bluecol">訓練計畫(在職)
                                </td>
                                <td class="whitecol">
                                    <asp:CheckBoxList ID="chkTPlanID1" runat="server" CellSpacing="0" CellPadding="0" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font">
                                    </asp:CheckBoxList>
                                    <input id="TPlanID1HID" value="0" type="hidden" name="TPlanID1HID" runat="server">
                                </td>
                            </tr>
                            <tr id="TPlanIDX_TR" runat="server">
                                <td class="bluecol">訓練計畫(其他)
                                </td>
                                <td class="whitecol">
                                    <asp:CheckBoxList ID="chkTPlanIDX" runat="server" CellSpacing="0" CellPadding="0" RepeatColumns="3" RepeatDirection="Horizontal" CssClass="font">
                                    </asp:CheckBoxList>
                                    <input id="TPlanIDXHID" value="0" type="hidden" name="TPlanIDXHID" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol" colspan="2" align="center">
                                    <asp:Button ID="btnSave1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                                <asp:Button ID="btnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
            <%--
				<TR>
					<TD>
					</TD>
				</TR>
            --%>
        </table>
        <input id="hidslid" type="hidden" runat="server">
    </form>
</body>
</html>
