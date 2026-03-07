 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SV_01_001.aspx.vb" Inherits="WDAIIP.SV_01_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>問卷名稱設定</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script src="../../js/common.js"></script>
	<script src="../../js/TIMS.js"></script>
	<script language="javascript">
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tr>
			<td>
				首頁&gt;&gt;系統管理&gt;&gt;問卷管理&gt;&gt;問卷名稱設定
			</td>
		</tr>
	</table>
	<asp:Panel ID="table_F" runat="server">
		<table class="table_nw" width="740" runat="server" cellpadding="1" cellspacing="1">
			<tr>
				<td class="bluecol" width="100">
					<label>
						問卷名稱</label>
				</td>
				<td class="whitecol">
					<input id="Ipt_Name" style="width: 373px; height: 22px" type="text" maxlength="100" size="56" runat="server" height="18">
				</td>
			</tr>
			<tr id="TRddlSType" runat="server">
				<td class="bluecol" width="100">
					<label>
						問卷種類</label>
				</td>
				<td class="whitecol">
					<asp:DropDownList ID="ddlSType" runat="server">
					</asp:DropDownList>
				</td>
			</tr>
		</table>
		<table width="740">
			<tr align="center">
				<td class="whitecol">
					<asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="button_b_S" />
					<asp:Button ID="btnInsert" runat="server" Text="新增" CssClass="button_b_S" />
				</td>
			</tr>
		</table>
	</asp:Panel>
	<table width="740" runat="server">
		<tr width="100%">
			<td align="center" width="100%">
				<font face="新細明體">
					<asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" runat="server">
						<AlternatingItemStyle HorizontalAlign="Center" BackColor="#F5F5F5"></AlternatingItemStyle>
						<HeaderStyle CssClass="head_navy"></HeaderStyle>
						<Columns>
							<asp:BoundColumn HeaderText="序號">
								<HeaderStyle Width="5%"></HeaderStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="Name" HeaderText="問卷名稱">
								<HeaderStyle Width="70%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Left" VerticalAlign="Middle"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="Avail" HeaderText="狀態">
								<HeaderStyle Width="10%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
							</asp:BoundColumn>
							<asp:TemplateColumn HeaderText="功能">
								<HeaderStyle Width="20%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center" />
								<ItemTemplate>
									<asp:Button ID="Btn_edit" runat="server" Text="修改" CommandName="edit"></asp:Button>
									<asp:Button ID="Btn_del" runat="server" Text="刪除" CommandName="del"></asp:Button>
								</ItemTemplate>
							</asp:TemplateColumn>
						</Columns>
						<PagerStyle Visible="False"></PagerStyle>
					</asp:DataGrid></font><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
			</td>
		</tr>
		<tr width="100%">
			<td width="100%">
				<p align="center">
					<asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></p>
			</td>
		</tr>
	</table>
	<asp:Panel ID="table_I" runat="server">
		<table class="table_nw" width="740" cellpadding="1" cellspacing="1">
			<tr>
				<td class="bluecol" width="100">
					<label>
						問卷名稱</label>
				</td>
				<td class="whitecol">
					<input id="IputQName" style="width: 365px; height: 22px" type="text" maxlength="100" size="55" runat="server" height="18">
				</td>
			</tr>
			<tr id="TRddlSurveyType" runat="server">
				<td class="bluecol">
					<label>
						問卷種類</label>
				</td>
				<td class="whitecol" style="height: 18px">
					<asp:DropDownList ID="ddlSurveyType" runat="server">
					</asp:DropDownList>
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					<label>
						狀態</label>
				</td>
				<td class="whitecol">
					<input id="ISUSE" type="checkbox" checked value="" name="" runat="server">啟用
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					<label>
						非一般問卷</label>
				</td>
				<td class="whitecol">
					<asp:CheckBox ID="chkinternal" runat="server" Text="(供內部使用)" />
				</td>
			</tr>
		</table>
		<table width="740">
			<tr align="center" runat="server">
				<td class="whitecol">
					<asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_S" />
					<asp:Button ID="btnReturn1" runat="server" Text="回上一頁" CssClass="asp_button_S" />
					<%--	
					<input id="save" type="button" value="儲存" name="save" runat="server" class="asp_button_S"  >
					<input id="return1" type="button" value="回上一頁" name="return1" runat="server" class="button_b_S">
					--%>
				</td>
			</tr>
		</table>
	</asp:Panel>
	<asp:HiddenField ID="HidSVID" runat="server" />
	<asp:HiddenField ID="HidMode" runat="server" />
	</form>
</body>
</html>
