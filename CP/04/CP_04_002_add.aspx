<%@ Page Language="vb" EnableEventValidation="false" AutoEventWireup="true" CodeBehind="CP_04_002_add.aspx.vb" Inherits="WDAIIP.CP_04_002_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>CP_04_002_add</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/common.js"></script>
	<style type="text/css">
		.class_link A { color: #000000; }
		.class_link A:link { color: #0000ff; }
		.class_link A:hover { color: #0000ff; }
		A:visited { color: #0000ff; }
		A:active { color: #0000ff; }
	</style>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" width="100%">
		<tr>
			<td>
				<font class="font" size="2">首頁&gt;&gt;訓練查核與績效管理&gt;&gt;訓練資料查詢&gt;&gt;</font><font class="font" color="#800000" size="2">計畫資料</font>
			</td>
		</tr>
	</table>
	<table class="font" id="Table6" cellspacing="0" cellpadding="1" width="100%" border="0">
		<tbody>
			<tr>
				<td style="width: 10%">
					<font face="新細明體" size="2">
						<asp:Label ID="Year" runat="server" CssClass="font">年度：</asp:Label><asp:Label ID="YearLabel" runat="server"></asp:Label></font>
				</td>
				<td style="width: 35%">
					<font face="新細明體" size="2">
						<asp:Label ID="District" runat="server">轄區：</asp:Label><asp:Label ID="DistrictLabel" runat="server"></asp:Label></font>
				</td>
				<td style="width: 15%">
					<font face="新細明體" size="2">
						<asp:Label ID="Count" runat="server" CssClass="font">筆數：</asp:Label><asp:Label ID="CountLabel" runat="server"></asp:Label></font>
				</td>
				<td style="width: 20%">
					<asp:Label ID="Label2" runat="server" CssClass="font">訓練總人數：</asp:Label><asp:Label ID="STNum" runat="server"></asp:Label>
				</td>
				<td style="width: 20%">
					<asp:Label ID="Label4" runat="server" CssClass="font">訓練總經費：</asp:Label></FONT><asp:Label ID="SumTotalCost" runat="server"></asp:Label>
				</td>
			</tr>
		</tbody>
	</table>
	<table class="font" id="Table5" cellspacing="0" cellpadding="0" width="100%" border="0">
		<tr>
			<td>
				<div id="Div1" runat="server">
					<asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AutoGenerateColumns="False" AllowSorting="True" AllowPaging="True" DataKeyField="SeqNO" Width="100%" PageSize="20">
						<AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
						<HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
						<Columns>
							<asp:BoundColumn HeaderText="序號"></asp:BoundColumn>
							<asp:BoundColumn DataField="DistName" SortExpression="DistID" HeaderText="轄區">
								<HeaderStyle ForeColor="Blue"></HeaderStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="AppliedDate" HeaderText="申請日期" DataFormatString="{0:d}"></asp:BoundColumn>
							<asp:BoundColumn DataField="PlanName" SortExpression="PlanID" HeaderText="訓練計畫">
								<HeaderStyle ForeColor="Blue"></HeaderStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="訓練機構">
								<HeaderStyle ForeColor="Blue"></HeaderStyle>
							</asp:BoundColumn>
							<asp:ButtonColumn DataTextField="ClassName" HeaderText="班名" CommandName="SeqNO">
								<ItemStyle CssClass="class_link"></ItemStyle>
							</asp:ButtonColumn>
							<asp:BoundColumn DataField="AppliedResult" HeaderText="審核狀態"></asp:BoundColumn>
							<asp:BoundColumn DataField="STDate" SortExpression="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
								<HeaderStyle ForeColor="Blue"></HeaderStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="FDDate" HeaderText="結訓日期" DataFormatString="{0:d}"></asp:BoundColumn>
							<asp:BoundColumn DataField="TrainNum" HeaderText="招生人數"></asp:BoundColumn>
							<asp:BoundColumn DataField="THours" HeaderText="時數"></asp:BoundColumn>
							<asp:BoundColumn HeaderText="訓練費用"></asp:BoundColumn>
							<asp:BoundColumn Visible="False" DataField="PlanID" HeaderText="計畫代碼"></asp:BoundColumn>
							<asp:BoundColumn Visible="False" DataField="ComIDNO" HeaderText="廠商統一編號"></asp:BoundColumn>
							<asp:BoundColumn Visible="False" DataField="TrainName" HeaderText="訓練職類"></asp:BoundColumn>
							<asp:TemplateColumn Visible="False" HeaderText="功能">
								<HeaderStyle Width="100px"></HeaderStyle>
								<ItemTemplate>
									<asp:Button runat="server" Text="詳細" CommandName="Company_Seqno" CausesValidation="false" ID="Button1"></asp:Button>
								</ItemTemplate>
							</asp:TemplateColumn>
						</Columns>
						<PagerStyle Visible="False"></PagerStyle>
					</asp:DataGrid>
				</div>
				<div align="center">
					<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
				</div>
			</td>
		</tr>
		<tr>
			<td style="height: 16px" align="center">
				<asp:Label ID="NoData" runat="server" CssClass="font"></asp:Label>
			</td>
		</tr>
		<tr>
			<td align="center">
				<asp:Button ID="Button2" runat="server" CausesValidation="False" Text="回上頁" CssClass="asp_button_S"></asp:Button>&nbsp;
				<asp:Button Style="z-index: 0" ID="btnExport1" runat="server" Text="匯出Excel" CssClass="asp_Export_M"></asp:Button>
			</td>
		</tr>
		<tr>
			<td align="left">
				<asp:Label ID="description" runat="server" CssClass="font">排序說明：以轄區、訓練計畫、開訓日期做排序</asp:Label>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
