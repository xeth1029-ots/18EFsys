 

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_011.aspx.vb" Inherits="WDAIIP.SD_01_011" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
	<title>報名學員補助金歷史查詢</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
	<meta content="JavaScript" name="vs_defaultClientScript" />
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
	<link href="../../css/style.css" type="text/css" rel="stylesheet" />
	<script type="text/javascript" src="../../js/date-picker.js"></script>
	<script type="text/javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" src="../../js/common.js"></script>
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
			if (document.getElementById('OCID1').value == '')
			{ document.getElementById('Button7').click(); }
			openClass('../02/SD_02_ch.aspx?RID=' + document.getElementById('RIDValue').value);
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="Frametable" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;<font color="#990000">報名學員補助金歷史查詢</font>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table id="Page1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
					<tr>
						<td>
							<table class="table_nw" id="table2" cellspacing="1" cellpadding="1" width="100%">
								<tr id="Orgtr" runat="server">
									<td class="bluecol" width="100">
										訓練機構
									</td>
									<td class="whitecol" colspan="3">
										<asp:TextBox ID="center" runat="server" Width="410px"></asp:TextBox>
										<input id="RIDValue" type="hidden" name="Hidden2" runat="server" />
										<input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
										<asp:Button ID="Button7" Style="display: none" runat="server"></asp:Button>
										<asp:Button ID="Button6" Style="display: none" runat="server" Text="Button6"></asp:Button>
										<span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
											<asp:Table ID="HistoryRID" runat="server" Width="310px">
											</asp:Table>
										</span>
									</td>
								</tr>
								<tr>
									<td class="bluecol">
										職類/班別
									</td>
									<td class="whitecol" colspan="3">
										<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
										<asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
										<input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
										<input id="Button4" type="button" value="清除" name="Button4" runat="server" class="asp_button_S" />
										<input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
										<input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
										<span id="HistoryList" style="position: absolute; display: none; left: 270px">
											<asp:Table ID="Historytable" runat="server" Width="310">
											</asp:Table>
										</span>
									</td>
								</tr>
								<tr>
									<td class="bluecol">
										身分證號碼
									</td>
									<td class="whitecol">
										<asp:TextBox ID="IDNO" runat="server"></asp:TextBox>
									</td>
									<td class="bluecol" width="100">
										學員姓名
									</td>
									<td class="whitecol">
										<asp:TextBox ID="Name" runat="server"></asp:TextBox>
									</td>
								</tr>
							</table>
							<table width="100%">
								<tr>
									<td class="whitecol" align="center" colspan="4">
										<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
										<asp:TextBox ID="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:TextBox>
										<asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
									</td>
								</tr>
								<tr>
									<td class="whitecol" align="center" colspan="4">
										<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
									</td>
								</tr>
							</table>
							<table id="DataGridtable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
								<tr>
									<td>
										<asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="true" CssClass="font" AutoGenerateColumns="False">
											<AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
											<HeaderStyle CssClass="head_navy" />
											<Columns>
												<asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
													<HeaderStyle Width="120px"></HeaderStyle>
												</asp:BoundColumn>
												<asp:BoundColumn DataField="Name" HeaderText="姓名">
													<HeaderStyle Width="60px"></HeaderStyle>
												</asp:BoundColumn>
												<asp:BoundColumn DataField="Birthday" HeaderText="出生日期" DataFormatString="{0:d}">
													<HeaderStyle Width="60px"></HeaderStyle>
												</asp:BoundColumn>
												<asp:TemplateColumn HeaderText="功能">
													<ItemTemplate>
														<asp:LinkButton ID="Button3" runat="server" Text="檢視" CssClass="linkbutton"></asp:LinkButton>
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
				<table class="table_nw" id="Page2" cellspacing="1" cellpadding="1" width="100%" runat="server">
					<tr>
						<td align="center" class="whitecol">
							<table class="font" id="table3" cellspacing="1" cellpadding="1" width="100%" border="0">
								<tr>
									<td class="bluecol" width="100">
										姓名
									</td>
									<td class="whitecol" width="200">
										<asp:Label ID="LName" runat="server"></asp:Label>
									</td>
									<td class="bluecol" width="100">
										身分證號碼
									</td>
									<td class="whitecol" width="200">
										<asp:Label ID="LIDNO" runat="server"></asp:Label>
									</td>
								</tr>
								<tr>
									<td colspan="4" class="whitecol">
										<asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
											<AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
											<HeaderStyle CssClass="head_navy" />
											<Columns>
												<asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
												<asp:BoundColumn DataField="ClassCName" HeaderText="參訓課程"></asp:BoundColumn>
												<asp:BoundColumn DataField="Stdate" HeaderText="開訓日期" DataFormatString="{0:d}">
													<HeaderStyle Width="80px"></HeaderStyle>
												</asp:BoundColumn>
												<asp:BoundColumn DataField="Ftdate" HeaderText="結訓日期" DataFormatString="{0:d}">
													<HeaderStyle Width="80px"></HeaderStyle>
												</asp:BoundColumn>
												<asp:BoundColumn DataField="SumOfMoney" HeaderText="申請補助金額">
													<HeaderStyle Width="100px"></HeaderStyle>
												</asp:BoundColumn>
												<asp:BoundColumn DataField="AppliedStatusM" HeaderText="審核狀態">
													<HeaderStyle Width="80px"></HeaderStyle>
												</asp:BoundColumn>
												<asp:BoundColumn DataField="AppliedStatus" HeaderText="撥款狀態">
													<HeaderStyle Width="80px"></HeaderStyle>
												</asp:BoundColumn>
											</Columns>
										</asp:DataGrid><asp:Label ID="msg2" runat="server"></asp:Label>
									</td>
								</tr>
								<tr>
									<%--撥款通過總額--%>
									<td colspan="4" class="whitecol">
										(補助總額：
										<asp:Label ID="LabTotal" runat="server"></asp:Label>)-(經費審核確認總額：
										<asp:Label ID="LabSumOfMoney" runat="server"></asp:Label>)=(剩餘可用額度：
										<asp:Label ID="RemainSub" runat="server"></asp:Label>)
									</td>
								</tr>
							</table>
					</tr>
				</table>
				<table width="100%">
					<tr>
						<td align="center">
							<asp:Button ID="Button5" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
