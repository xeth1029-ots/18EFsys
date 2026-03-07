<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_011.aspx.vb" Inherits="WDAIIP.SD_05_011" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_05_011</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script language="javascript">
        /*
		function but_del(stdid, id) {
			//if (is_parent) {
			//	alert("此班別代碼資料尚有開班資料檔參照,不可刪除!!");
			//return;
			//}
			if (window.confirm("此動作會刪除此筆歷史資料，是否確定刪除?"))
				location.href = 'SD_04_004_del.aspx?stdid=' + stdid + '&ID=' + id;
		}
        */
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;90-93歷史資料</asp:Label>
                </td>
            </tr>
        </table>
        <%--	<input id="check_mod" type="hidden" name="check_mod" runat="server">
	<input id="check_del" type="hidden" name="check_del" runat="server">--%>
        <%--<p><font face="新細明體">--%>
        <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
						<tr>
							<td>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">90-93歷史資料 </font></td>
						</tr>
					</table>--%>
                    <div align="left">

                        <table id="SearchTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tbody>
                                <tr>
                                    <td align="center">
                                        <div align="left">
                                            <table class="table_nw" id="Table3" width="100%" cellpadding="1" cellspacing="1">
                                                <tr>
                                                    <td class="bluecol" style="width: 20%">身分證號碼 </td>
                                                    <td class="whitecol" style="width: 30%">
                                                        <asp:TextBox ID="SID" runat="server" Width="40%"></asp:TextBox>
                                                    </td>
                                                    <td class="bluecol" style="width: 20%">姓名 </td>
                                                    <td class="whitecol" style="width: 30%">
                                                        <asp:TextBox ID="Name" runat="server" Width="40%"></asp:TextBox>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol_need" style="width: 20%">年度 </td>
                                                    <td class="whitecol" style="width: 30%">
                                                        <asp:DropDownList ID="YearList" runat="server">
                                                            <asp:ListItem Value="0">===請選擇===</asp:ListItem>
                                                            <asp:ListItem Value="2001">2001</asp:ListItem>
                                                            <asp:ListItem Value="2002">2002</asp:ListItem>
                                                            <asp:ListItem Value="2003">2003</asp:ListItem>
                                                            <asp:ListItem Value="2004">2004</asp:ListItem>
                                                        </asp:DropDownList>
                                                        <%--<asp:RequiredFieldValidator ID="Re_Years_List" runat="server" InitialValue="0" ErrorMessage="請選擇年度" Display="None" ControlToValidate="YearList"></asp:RequiredFieldValidator>--%>
                                                    </td>
                                                    <%--<td class="bluecol" style="width: 20%">轄區中心 </td>--%>
                                                    <td class="bluecol" style="width: 20%">轄區分署 </td>
                                                    <td class="whitecol" style="width: 30%">
                                                        <asp:DropDownList ID="DistrictList" runat="server" Width="160px">
                                                        </asp:DropDownList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="bluecol">訓練計畫 </td>
                                                    <td colspan="3" class="whitecol">
                                                        <asp:DropDownList ID="Plan_List" runat="server">
                                                        </asp:DropDownList>
                                                    </td>
                                                </tr>
                                            </table>
                                            <table width="100%">
                                                <tr>
                                                    <td align="center" class="whitecol">
                                                        <p>
                                                            <asp:Button ID="Button2" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                                            <asp:Button ID="Button1" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                                        </p>
                                                        <p>
                                                            <%--<asp:ValidationSummary ID="Summary" runat="server" Width="440px" ShowMessageBox="True" ShowSummary="False" DisplayMode="List"></asp:ValidationSummary>--%>
                                                            <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                                        </p>
                                                    </td>
                                                </tr>
                                            </table>
                                        </div>
                                        <div align="left">
                                            &nbsp;
                                        </div>
                                        <div align="left">
                                            <asp:Panel ID="Panel2" runat="server" Width="100%" Visible="False">
                                                <asp:DataGrid ID="Stud_DG" runat="server" CssClass="font" Width="100%" Visible="False" AutoGenerateColumns="False" AllowPaging="True" AllowCustomPaging="True" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <%--<asp:BoundColumn DataField="DistName" HeaderText="轄區中心" HeaderStyle-Width="10%"></asp:BoundColumn>--%>
                                                        <asp:BoundColumn DataField="DistName" HeaderText="轄區分署" HeaderStyle-Width="10%"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫" HeaderStyle-Width="15%"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="TrinUnit" HeaderText="培訓單位" HeaderStyle-Width="15%"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="ClassName" HeaderText="班別" HeaderStyle-Width="10%"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="Name" HeaderText="姓名" HeaderStyle-Width="10%" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="Sex" HeaderText="性別" HeaderStyle-Width="10%"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="IDNO" HeaderText="身分證" HeaderStyle-Width="10%"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="Ident" HeaderText="身分別" HeaderStyle-Width="10%"></asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="10%">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Button ID="edit_but" runat="server" Text="修改" CommandName="edit" class="asp_button_M"></asp:Button>
                                                                <asp:Button ID="del_but" runat="server" Text="刪除" CommandName="del" class="asp_button_M"></asp:Button>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:BoundColumn Visible="False" DataField="source" HeaderText="來源"></asp:BoundColumn>
                                                        <asp:BoundColumn Visible="False" DataField="Stdid" HeaderText="序號">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn Visible="False" DataField="DistID" HeaderText="DistID"></asp:BoundColumn>
                                                        <asp:BoundColumn Visible="False" DataField="TPlanID" HeaderText="TPlanID"></asp:BoundColumn>
                                                    </Columns>
                                                    <PagerStyle Visible="False"></PagerStyle>
                                                </asp:DataGrid>
                                                <div align="center">
                                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                                </div>
                                            </asp:Panel>
                                        </div>
                                    </td>
                                </tr>
                            </tbody>
                        </table>

                    </div>
                </td>
            </tr>
        </table>
        <%--</font></p>--%>
    </form>
</body>
</html>
