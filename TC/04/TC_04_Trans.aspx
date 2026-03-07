<%@ Page Language="vb" AutoEventWireup="false" Codebehind="TC_04_Trans.aspx.vb" Inherits="TIMS.TC_04_Trans" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>訓練機構管理</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../style.css" type="text/css" rel="stylesheet">
		<script language="javascript" src="../../js/openwin/openwin.js"></script>
	</HEAD>
	<body>
		<form id="form1" method="post" runat="server">
			<input id="RIDValue" type="hidden" name="RIDValue" runat="server"> <input id="PlanIDValue" type="hidden" name="PlanIDValue" runat="server">
			<table class="font" cellSpacing="1" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="font" width="100%">首頁&gt;&gt;訓練機構管理&gt;&gt;<font color="#990000">計畫審核作業</font></td>
				</tr>
				<tr>
					<td>
						<table class="font" cellSpacing="0" borderColorDark="#ffffff" cellPadding="1" width="75%"
							borderColorLight="#666666" border="1">
							<tr>
								<td style="HEIGHT: 17px" width="20%" bgColor="#ffcccc">&nbsp;&nbsp;&nbsp; 階層</td>
								<td style="HEIGHT: 17px" width="50%" bgColor="#fff9e1" colSpan="3">&nbsp;&nbsp;&nbsp;<asp:dropdownlist id="OrgLevel" runat="server" AutoPostBack="True">
										<asp:ListItem>===請選擇===</asp:ListItem>
										<asp:ListItem Value="1">中心</asp:ListItem>
										<asp:ListItem Value="2">下層單位</asp:ListItem>
									</asp:dropdownlist></td>
							</tr>
							<tr>
								<td bgColor="#ffcccc">&nbsp;&nbsp;&nbsp; 計劃階層</td>
								<td bgColor="#fff9e1" colSpan="3">&nbsp;&nbsp;&nbsp;<asp:textbox id="TBplan" runat="server" Columns="30"></asp:textbox><input id="choice_button" onclick="javascript:wopen('../../Common/LevPlan.aspx','計畫階段',500,500,1)"
										type="button" value="選擇" name="choice_button" runat="server">
								</td>
							</tr>
							<tr>
								<td bgColor="#ffcccc">&nbsp;&nbsp;&nbsp; 轄區中心</td>
								<td bgColor="#fff9e1" colSpan="3">&nbsp;&nbsp;&nbsp;<asp:dropdownlist id="DistID" runat="server"></asp:dropdownlist></td>
							</tr>
						</table>
						<table width="75%" border="0">
							<tr>
								<td width="75%">
									<div align="center"><asp:button id="btnAdd" runat="server" Text="儲存"></asp:button></div>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
