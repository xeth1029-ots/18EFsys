<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_08_007.aspx.vb" Inherits="WDAIIP.SD_08_007" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>申請查詢</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script src="../../js/common.js"></script>
	<script language="javascript">
		function chkdata() {
			var form = document.form1;
			var msg = '';

			if (!checkDate(form.SCHASDate.value) && form.SCHASDate.value != '') msg += '申請日期(起)格式不正確\n';

			if (!checkDate(form.SCHAEDate.value) && form.SCHAEDate.value != '') msg += '申請日期(迄)格式不正確\n';

			if (form.SCHSAge.value != '' && !isInt(form.SCHSAge.value)) msg += '年齡(起)格式不正確\n';

			if (form.SCHEAge.value != '' && !isInt(form.SCHEAge.value)) msg += '年齡(迄)格式不正確\n';

			if (form.SCHSMonth.value != '' && !(isInt(form.SCHSMonth.value) || isFloat(form.SCHSMonth.value))) msg += '核發月數(起)格式不正確\n';

			if (form.SCHEMonth.value != '' && !(isInt(form.SCHEMonth.value) || isFloat(form.SCHEMonth.value))) msg += '核發月數(迄)格式不正確\n';

			if (form.SCHSMoney.value != '' && !isInt(form.SCHSMoney.value)) msg += '核發金額(起)格式不正確\n';

			if (form.SCHEMoney.value != '' && !isInt(form.SCHEMoney.value)) msg += '核發金額(迄)格式不正確\n';


			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		function cleardata() {

			var form = document.form1;

			form.center.value = '';

			form.RIDValue.value = '';

			form.orgid_value.value = '';

			form.SCHIdentityID.value = '';

			form.SCHIDNO.value = '';

			form.SCHSex1.checked = false;

			form.SCHSex2.checked = false;

			form.SCHASDate.value = '';

			form.SCHAEDate.value = '';

			form.SCHSAge.value = '';

			form.SCHEAge.value = '';

			form.SCHSMonth.value = '';

			form.SCHEMonth.value = '';

			form.SCHSMoney.value = '';

			form.SCHEMoney.value = '';
		}

	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="Table2" cellspacing="1" width="100%" border="0">
					<tr>
						<td>首頁&gt;&gt;學員動態管理&gt;&gt;職業訓練生活津貼&gt;&gt;<font color="#990000">申請查詢</font> </td>
					</tr>
				</table>
				<asp:Panel ID="searchPanel" runat="server">
					<table id="SearchTable" class="table_nw" cellspacing="1" cellpadding="1" width="100%">
						<tr>
							<td class="bluecol">訓練機構 </td>
							<td class="whitecol" colspan="3">
								<asp:TextBox ID="center" runat="server" Width="310px" onfocus="this.blur()"></asp:TextBox>
								<input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
								<input id="orgid_value" type="hidden" name="orgid_value" runat="server" />
								<input id="Button7" value="..." type="button" name="Button7" runat="server" class="asp_button_Mini" />
							</td>
						</tr>
						<tr>
							<td class="bluecol">身分別 </td>
							<td class="whitecol" colspan="3">
								<asp:DropDownList ID="SCHIdentityID" runat="server">
								</asp:DropDownList>
							</td>
						</tr>
						<tr>
							<td class="bluecol" width="100">身分證號 </td>
							<td class="whitecol" width="270">
								<asp:TextBox ID="SCHIDNO" runat="server" Columns="15"></asp:TextBox>
							</td>
							<td class="bluecol" width="100">性 別 </td>
							<td class="whitecol" width="270">
								<input id="SCHSex1" type="checkbox" runat="server">男 &nbsp;&nbsp;&nbsp;
								<input id="SCHSex2" type="checkbox" runat="server">女 </td>
						</tr>
						<tr>
							<td class="bluecol">申請日期 </td>
							<td class="whitecol">
								<asp:TextBox ID="SCHASDate" runat="server" Columns="7"></asp:TextBox><img style="cursor: pointer" id="IMG1" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" runat="server" />
								～
								<asp:TextBox ID="SCHAEDate" runat="server" Columns="7"></asp:TextBox><img style="cursor: pointer" id="IMG2" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" runat="server" />
							</td>
							<td class="bluecol">年 齡 </td>
							<td class="whitecol">
								<asp:TextBox ID="SCHSAge" runat="server" Columns="4"></asp:TextBox>～
								<asp:TextBox ID="SCHEAge" runat="server" Columns="4"></asp:TextBox>
							</td>
						</tr>
						<tr>
							<td class="bluecol">申請月數 </td>
							<td class="whitecol">
								<asp:TextBox ID="SCHSMonth" runat="server" Columns="4"></asp:TextBox>～
								<asp:TextBox ID="SCHEMonth" runat="server" Columns="4"></asp:TextBox>
							</td>
							<td class="bluecol">申請金額 </td>
							<td class="whitecol">
								<asp:TextBox ID="SCHSMoney" runat="server" Columns="6"></asp:TextBox>～
								<asp:TextBox ID="SCHEMoney" runat="server" Columns="6"></asp:TextBox>
							</td>
						</tr>
					</table>
					<table width="100%">
						<tr>
							<td align="center" class="whitecol">
								<asp:Button ID="Search" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;&nbsp;&nbsp;
								<input onclick="cleardata();" value="重設" type="button" class="asp_button_S" />
							</td>
						</tr>
					</table>
					<table id="ShowDataTable" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
						<tr>
							<td align="center">
								<asp:Label ID="mesg" runat="server" ForeColor="#ff0000"></asp:Label>
							</td>
						</tr>
						<tr>
							<td>
								<asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowSorting="True" AllowPaging="True" AutoGenerateColumns="False" CssClass="font">
									<ItemStyle BackColor="White"></ItemStyle>
									<AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
									<HeaderStyle CssClass="head_navy" />
									<Columns>
										<asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="單位名稱" ItemStyle-HorizontalAlign="Left">
											<HeaderStyle></HeaderStyle>
										</asp:BoundColumn>
										<asp:TemplateColumn HeaderText="身分證號<br>姓名">
											<HeaderStyle Width="12%"></HeaderStyle>
											<ItemTemplate>
												<asp:LinkButton runat="server" ID="txtidno" ForeColor="#0000FF" CommandName="View"></asp:LinkButton>
												<br>
												<asp:Label runat="server" ID="txtname"></asp:Label>
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="出生日期<br>申請日期">
											<HeaderStyle Width="10%"></HeaderStyle>
											<ItemTemplate>
												<asp:Label runat="server" ID="birth"></asp:Label>
												<br>
												<asp:Label runat="server" ID="apply"></asp:Label>
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:BoundColumn HeaderText="受訓起訖">
											<HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
										</asp:BoundColumn>
										<asp:BoundColumn HeaderText="申請月數<br>金額">
											<HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
										</asp:BoundColumn>
										<asp:TemplateColumn HeaderText="初審" HeaderStyle-Width="8%">
											<ItemTemplate>
												<asp:Label ID="asf" runat="server"></asp:Label>
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="送勞<br>保局" HeaderStyle-Width="8%">
											<ItemTemplate>
												<asp:Label ID="isdl" runat="server"></asp:Label>
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="勾稽" HeaderStyle-Width="8%">
											<ItemTemplate>
												<asp:Label ID="asfin" runat="server"></asp:Label>
											</ItemTemplate>
										</asp:TemplateColumn>
									</Columns>
								</asp:DataGrid>
							</td>
						</tr>
					</table>
					<table id="TablePage" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
						<tr>
							<td>
								<p align="center">
									<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
								</p>
								<p align="center">&nbsp;</p>
							</td>
						</tr>
					</table>
				</asp:Panel>
				<asp:Panel ID="historyPanel" runat="server">
					<table id="ShowDataTable1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
						<tr>
							<td colspan="3" align="center">
								<p>
									<asp:Label ID="msg1" runat="server" ForeColor="#ff0000"></asp:Label></p>
							</td>
						</tr>
						<tr>
							<td width="35%"><b>姓 名：</b>
								<asp:Label ID="headname" runat="server"></asp:Label>
							</td>
							<td><b>身分證號：</b>
								<asp:Label ID="headidno" runat="server"></asp:Label>
							</td>
							<td width="35%"><b>生 日：</b>
								<asp:Label ID="headbrith" runat="server"></asp:Label>
							</td>
						</tr>
						<tr>
							<td colspan="3">
								<asp:DataGrid ID="Datagrid2" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font">
									<ItemStyle BackColor="White"></ItemStyle>
									<AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
									<HeaderStyle CssClass="head_navy" />
									<Columns>
										<asp:BoundColumn HeaderText="序號">
											<HeaderStyle Width="5%"></HeaderStyle>
										</asp:BoundColumn>
										<asp:BoundColumn DataField="applydate" HeaderText="申請日期" DataFormatString="{0:d}">
											<HeaderStyle Width="10%"></HeaderStyle>
										</asp:BoundColumn>
										<asp:BoundColumn HeaderText="受訓起訖">
											<HeaderStyle Width="10%"></HeaderStyle>
										</asp:BoundColumn>
										<asp:BoundColumn HeaderText="申請月數&lt;br&gt;金額">
											<HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
										</asp:BoundColumn>
										<asp:BoundColumn HeaderText="實領月數&lt;br&gt;金額">
											<HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
										</asp:BoundColumn>
										<asp:TemplateColumn HeaderText="初審">
											<HeaderStyle Width="8%"></HeaderStyle>
											<ItemTemplate>
												<asp:Label ID="asf1" runat="server"></asp:Label>
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="送勞&lt;br&gt;保局">
											<HeaderStyle Width="8%"></HeaderStyle>
											<ItemTemplate>
												<asp:Label ID="isdl1" runat="server"></asp:Label>
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="勾稽">
											<HeaderStyle Width="8%"></HeaderStyle>
											<ItemTemplate>
												<asp:Label ID="asfin1" runat="server"></asp:Label>
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:BoundColumn DataField="FailReasonFin" HeaderText="勾稽備註"></asp:BoundColumn>
									</Columns>
									<PagerStyle Visible="False"></PagerStyle>
								</asp:DataGrid>
							</td>
						</tr>
						<tr>
							<td colspan="3" align="center">
								<asp:Button ID="BackButton" runat="server" Width="50px" Text="回上頁" CssClass="asp_button_S"></asp:Button>
							</td>
						</tr>
					</table>
					<table id="TablePage2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
						<tr>
							<td>
								<p align="center">&nbsp;</p>
							</td>
						</tr>
					</table>
				</asp:Panel>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
