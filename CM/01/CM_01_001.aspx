 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_01_001.aspx.vb" Inherits="WDAIIP.CM_01_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>CM_01_001</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" src="../../js/common.js"></script>
	<script type="text/javascript" src="../../js/TIMS.js"></script>
	<script type="text/javascript" language="javascript">
		function choose_class() {
			wopen('../../SD/02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value, 'Class', 540, 520, 1);
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td class="font">
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
				            首頁&gt;&gt;訓練經費控管&gt;&gt;<font color="#990000">訓練計畫核銷作業</font>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
					<tr>
						<td id="td6" width="100" runat="server" class="bluecol">
							機構
						</td>
						<td colspan="5" rowspan="1" class="whitecol">
							<font face="新細明體">
								<asp:TextBox ID="center" runat="server" Width="410px" onfocus="this.blur()"></asp:TextBox>
								<input id="Button1" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
								<input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server">
								<input id="Re_ID" style="width: 32px; height: 22px" type="hidden" name="Re_ID" runat="server">
								<span id="HistoryList2" style="display: none; position: absolute">
									<asp:Table ID="HistoryRID" runat="server" Width="310px">
									</asp:Table>
								</span></font>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							職類/班別
						</td>
						<td colspan="5" class="whitecol">
							<font face="新細明體"><font face="新細明體"></font>
								<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
								<asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
								<input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
								<input id="OCIDValue1" style="width: 48px; height: 22px" type="hidden" size="2" runat="server">
								<input id="TMIDValue1" style="width: 40px; height: 22px" type="hidden" runat="server">
								<span id="HistoryList" style="display: none; left: 265px; position: absolute">
									<asp:Table ID="HistoryTable" runat="server" Width="310">
									</asp:Table>
								</span></font>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							開訓日期區間
						</td>
						<td colspan="5" class="whitecol">
							<font face="新細明體">
								<asp:TextBox ID="start_date" Width="80" onfocus="this.blur()" runat="server"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ～
								<asp:TextBox ID="end_date" Width="80" onfocus="this.blur()" runat="server"></asp:TextBox>
                                <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
							</font>
                            <font face="新細明體">&nbsp;&nbsp;&nbsp;</font><font face="新細明體"><font face="新細明體">&nbsp;&nbsp;&nbsp;&nbsp;</font>&nbsp;&nbsp;</font>
						</td>
					</tr>
					<tr>
						<td id="td5" runat="server" class="bluecol">
							核銷結餘
						</td>
						<td width="200" class="whitecol">
							<font face="新細明體"></font>
							<asp:DropDownList ID="DropDownList1" runat="server">
								<asp:ListItem Value="0" Selected="True">請選擇</asp:ListItem>
								<asp:ListItem Value="1">未結清</asp:ListItem>
								<asp:ListItem Value="2">已結清</asp:ListItem>
								<asp:ListItem Value="3">超支</asp:ListItem>
							</asp:DropDownList>
						</td>
						<td width="100" class="bluecol">
							期別
						</td>
						<td width="200" colspan="3" class="whitecol">
							<font face="新細明體">
								<asp:TextBox ID="TB_cycltype" runat="server" Width="48px"></asp:TextBox></font>
						</td>
					</tr>
				</table>
				<table width="100%">
					<tr>
						<td align="center" class="whitecol">
							<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label><asp:TextBox ID="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:TextBox>
							<asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button><br>
							<asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label><br>
							附註:本系統僅供業務管理承辦人記錄經費運用情形掌控之用，確實經費核銷仍依會計系統為準。
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	<table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
		<tr>
			<td>
				<asp:DataGrid ID="DG_Budget" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" Visible="False">
					<AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
					<HeaderStyle CssClass="head_navy"></HeaderStyle>
					<Columns>
						<asp:BoundColumn HeaderText="序號">
							<HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="OrgName" HeaderText="機構">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ClassCName" SortExpression="ClassCName" HeaderText="班別名稱">
							<HeaderStyle HorizontalAlign="Center" ForeColor="#00ffff"></HeaderStyle>
							<ItemStyle HorizontalAlign="Left"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn SortExpression="STDate" HeaderText="開結訓日" DataFormatString="{0:d}">
							<HeaderStyle HorizontalAlign="Center" ForeColor="#00ffff"></HeaderStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="THours" HeaderText="訓練&lt;br&gt;時數">
							<HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn HeaderText="計畫人數/每人費用">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn Visible="False" HeaderText="就安人數/金額">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn Visible="False" HeaderText="就保人數/金額">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn Visible="False" HeaderText="公務人數/金額">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn HeaderText="計畫總經費">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn HeaderText="已核銷總金額">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn SortExpression="TotalAdmCancelCost" HeaderText="結餘總金額">
							<HeaderStyle HorizontalAlign="Center" ForeColor="#00ffff"></HeaderStyle>
							<ItemStyle HorizontalAlign="Right"></ItemStyle>
						</asp:BoundColumn>
						<asp:TemplateColumn HeaderText="功能">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemTemplate>
								<asp:Button ID="edit_but" runat="server" Text="核銷" CommandName="edit"></asp:Button>
							</ItemTemplate>
						</asp:TemplateColumn>
					</Columns>
					<PagerStyle Visible="False"></PagerStyle>
				</asp:DataGrid>
			</td>
		</tr>
		<tr>
			<td>
				<div align="center">
					<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
				</div>
			</td>
		</tr>
	</table>
	<input id="hidClose1" type="hidden" runat="server" size="1">
	<input type="hidden" id="hidClose2" runat="server" size="1">
	<input type="hidden" id="hidClose3" runat="server" size="1">
	<input type="hidden" id="hidClose4" runat="server" size="1">
	<input type="hidden" id="hidClose5" runat="server" size="1">
	<input type="hidden" id="hidClose6" runat="server" size="1">
	</form>
</body>
</html>
