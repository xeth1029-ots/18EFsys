<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_005.aspx.vb" Inherits="WDAIIP.SYS_06_005" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>產業關鍵字設定</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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

</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;<font color="#990000">產業關鍵字設定</font>
                </td>
            </tr>
        </table>

        <table id="FrameTable2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>

                    <asp:Panel ID="panelSearch" runat="server">
                        <table id="Table3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <%--
								<TR id="KID_6_TR" runat="server">
									<TD class="SYS_TD1" width="100">&nbsp;&nbsp;新興產業</TD>
									<TD class="SYS_TD2">
										<asp:checkboxlist id="KID_6" runat="server" RepeatColumns="5" Font-Size="X-Small" RepeatDirection="Horizontal"></asp:checkboxlist><INPUT id="KID_6_hid" value="0" type="hidden" name="HID_DepID_6" runat="server">
									</TD>
								</TR>
								<TR id="KID_10_TR" runat="server">
									<TD class="SYS_TD1">&nbsp;&nbsp;重點服務業</TD>
									<TD class="SYS_TD2">
										<asp:checkboxlist id="KID_10" runat="server" RepeatColumns="4" Font-Size="X-Small" RepeatDirection="Horizontal"></asp:checkboxlist><INPUT id="KID_10_hid" value="0" type="hidden" name="HID_DepID_6" runat="server">
									</TD>
								</TR>
								<TR id="KID_4_TR" runat="server">
									<TD class="SYS_TD1">&nbsp;&nbsp;新興智慧型產業</TD>
									<TD class="SYS_TD2">
										<asp:checkboxlist id="KID_4" runat="server" RepeatColumns="5" Font-Size="X-Small" RepeatDirection="Horizontal"></asp:checkboxlist><INPUT id="KID_4_hid" value="0" type="hidden" name="HID_DepID_6" runat="server">
									</TD>
								</TR>
                            --%>
                            <tr>
                                <td class="bluecol" style="width: 20%">關鍵字</td>
                                <td class="whitecol">
                                    <asp:TextBox Style="z-index: 0" ID="tKeyNAME_s" runat="server" MaxLength="15" Width="30%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">產業名稱關鍵字</td>
                                <td class="whitecol">
                                    <asp:TextBox Style="z-index: 0" ID="tKNAME_s" runat="server" MaxLength="15" Width="30%"></asp:TextBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">產業類別</td>
                                <td class="whitecol">
                                    <asp:RadioButtonList Style="z-index: 0" ID="rblDep_s" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="A" Selected="True">全部</asp:ListItem>
                                        <asp:ListItem Value="07">六大新興產業</asp:ListItem>
                                        <asp:ListItem Value="08">十大重點服務業</asp:ListItem>
                                        <asp:ListItem Value="06">四大新興智慧型產業</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="btnAddnew" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center">
                                    <asp:Button ID="btnExport1" runat="server" Text="匯出EXCEL" CssClass="asp_Export_M"></asp:Button>
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
                                    <div id="Div1" runat="server">
                                        <asp:DataGrid Style="z-index: 0" ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True">
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <HeaderStyle CssClass="head_navy" />
                                            <Columns>
                                                <asp:BoundColumn HeaderText="序號">
                                                    <HeaderStyle Width="30px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="產業類別">
                                                    <HeaderStyle Width="200"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lbDNAME" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="產業名稱項目">
                                                    <HeaderStyle Width="300"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lbKNAME" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="關鍵字">
                                                    <HeaderStyle Width="200"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lbKeyNAME" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle Width="70px"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="btnEdit1" runat="server" Text="修改" CommandName="Edit1" CssClass="linkbutton"></asp:LinkButton>
                                                        <asp:LinkButton ID="btnDelete1" runat="server" Text="刪除" CommandName="Delete1" CssClass="linkbutton"></asp:LinkButton>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                            <PagerStyle Visible="False"></PagerStyle>
                                        </asp:DataGrid>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="PanelEdit1" runat="server">
                        <table id="Table4" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">刪除舊資料</td>
                                <td class="whitecol">
                                    <asp:CheckBox Style="z-index: 0" ID="cb_delKID" runat="server" Text="刪除該產業項用資料"></asp:CheckBox>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">六大新興產業</td>
                                <td class="whitecol">
                                    <asp:DropDownList Style="z-index: 0" ID="ddlKID06" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">十大重點服務業</td>
                                <td class="whitecol">
                                    <asp:DropDownList Style="z-index: 0" ID="ddlKID10" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">四大新興智慧型產業</td>
                                <td class="whitecol">
                                    <asp:DropDownList Style="z-index: 0" ID="ddlKID04" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">關鍵字</td>
                                <td class="whitecol">
                                    <asp:TextBox Style="z-index: 0" ID="tKeyNAME" runat="server" MaxLength="300" Columns="40"></asp:TextBox><br>
                                    (分隔請用半型逗號,可新增多筆資料,但只會修改一筆)
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2" align="center">
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
        <input id="hidKeysID" type="hidden" runat="server" name="hidKeysID" />
    </form>
</body>
</html>
