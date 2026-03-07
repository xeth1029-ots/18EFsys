 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SV_01_002.aspx.vb" Inherits="WDAIIP.SV_01_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>問卷分類標題設定</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script src="../../js/common.js"></script>
	<script language="javascript">

		function Check() {
			var msg = '';
			if (document.getElementById('Name3').value == '') {
				msg = '請輸入【分類標題】!\n';
			}
			if (document.getElementById('Serial').value == '') {
				msg += '請輸入【排序序號】!\n';
			}
			if (document.getElementById('Serial').value != '') {
				if (!isUnsignedInt(document.getElementById('Serial').value)) {
					msg += '【排序序號】必須為數字!\n';
				}
			}
			if (msg != "") {
				alert(msg);
			}
		}


	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				首頁&gt;&gt;系統管理&gt;&gt;問卷管理&gt;&gt;問卷分類標題設定
			</td>
		</tr>
	</table>
	<asp:Panel ID="table_F" runat="server" Width="740">
		<table class="table_nw" width="740" cellspacing="1" cellpadding="1">
			<tr>
				<td class="bluecol">
					<label>
						問卷名稱</label>
				</td>
				<td class="whitecol">
					<input id="Ipt_Name" style="width: 373px; height: 22px" maxlength="100" size="56" name="Ipt_Name" runat="server" height="18">
				</td>
			</tr>
		</table>
		<table width="100%">
			<tr align="center">
				<td class="whitecol">
					<input id="search" type="button" value="查詢" name="search" runat="server" class="button_b_S">
				</td>
			</tr>
		</table>
	</asp:Panel>
	<table id="Table2" width="740" runat="server">
		<tr width="100%">
			<td align="center" width="100%">
				<font face="新細明體">
					<asp:DataGrid ID="DataGrid1" runat="server" Width="100%" runat="server" CssClass="font" AutoGenerateColumns="False" AllowPaging="True">
						<AlternatingItemStyle HorizontalAlign="Center" BackColor="#F5F5F5"></AlternatingItemStyle>
						<HeaderStyle CssClass="head_navy"></HeaderStyle>
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
									<asp:Button ID="Btn_edit" runat="server" Text="設定問卷分類標題"></asp:Button>
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
					<asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></p>
			</td>
		</tr>
	</table>
	<asp:Panel ID="Table3" runat="server">
		<table class="table_nw" width="740" border="0" cellpadding="1" cellspacing="1">
			<tr>
				<td class="bluecol">
					<label>
						問卷名稱</label>
				</td>
				<td class="whitecol">
					<asp:Label ID="QName" runat="server"></asp:Label>
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					<label>
						分類標題</label>
				</td>
				<td class="whitecol">
					<input id="Name3" style="width: 392px; height: 22px" size="60" runat="server">
				</td>
			</tr>
			<tr>
				<td class="bluecol">
					<label>
						排序序號</label>
				</td>
				<td class="whitecol">
					<input id="Serial" style="width: 32px; height: 22px" runat="server" value="0">(0:自動給一排序序號)
				</td>
			</tr>
		</table>
		<table width="100%">
			<tr align="center">
				<td class="whitecol">
					<%--<input id="Save" type="button" value="新增(儲存)" name="Save" runat="server" class="button_b_M" onclick="return Save_onclick()">--%>
					<asp:Button ID="btnSave1" runat="server" Text="新增(儲存)" CssClass="button_b_M" />
					<input id="Return1" type="button" value="回上一頁" name="Return1" runat="server" class="button_b_M">
				</td>
			</tr>
		</table>
	</asp:Panel>
	<table id="Table4" width="740" runat="server">
		<tr width="100%">
			<td align="center" width="100%">
				<font face="新細明體">
					<asp:DataGrid ID="Datagrid2" runat="server" Width="100%" runat="server" CssClass="font" AutoGenerateColumns="False" AllowPaging="True">
						<AlternatingItemStyle HorizontalAlign="Center" BackColor="#F5F5F5"></AlternatingItemStyle>
						<HeaderStyle CssClass="head_navy"></HeaderStyle>
						<ItemStyle HorizontalAlign="Center" />
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
	<input id="SVID2" type="hidden" runat="server">
	<input id="MODE" type="hidden" runat="server">
	<input id="SKID2" type="hidden" runat="server">
	</form>
</body>
</html>
