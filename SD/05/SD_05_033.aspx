<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_033.aspx.vb" Inherits="WDAIIP.SD_05_033" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>參訓學員自行負擔費用清冊</title>
	<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
	<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
	<meta name="vs_defaultClientScript" content="JavaScript">
	<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	<link href="../../css/style.css" type="text/css" rel="stylesheet" />
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script language="javascript">
		function GETvalue() {
			document.getElementById('btnGETvalue1').click();
		}
		function SetOneOCID() {
			document.getElementById('btnSetOneOCID').click();
		}
		function choose_class() {
			var RIDValue = document.getElementById("RIDValue");
			openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
		}

		function search1() {
			var msg = '';
			var OCIDValue1 = document.getElementById("OCIDValue1");
			//if (OCIDValue1.value == '') {msg += '請選擇班級職類\n';}
			if (msg != '') {
				alert(msg);
				return false;
			}
		}
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tr>
			<td>
				<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">參訓學員自行負擔費用清冊</font>
							</asp:Label>
						</td>
					</tr>
				</table>
				<table id="tbSearch1" runat="server" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td class="bluecol" width="120">
							訓練機構
						</td>
						<td class="whitecol">
							<asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
							<input id="RIDValue" type="hidden" name="RIDValue" runat="server">
							<input id="btnSetLevOrg" type="button" value="..." name="btnSetLevOrg" runat="server" class="asp_button_Mini">
							<asp:Button ID="btnSetOneOCID" Style="display: none" runat="server"></asp:Button>
							<asp:Button ID="btnGETvalue1" Style="display: none" runat="server"></asp:Button>
							<span onclick="GETvalue()" id="HistoryList2" style="position: absolute; display: none">
								<asp:Table ID="HistoryRID" runat="server" Width="310px">
								</asp:Table>
							</span>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							職類/班別
						</td>
						<td class="whitecol">
							<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
							<asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
							<input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" />
							<input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" />
							<input onclick="choose_class()" type="button" value="..." class="asp_button_Mini">
							<span id="HistoryList" style="position: absolute; left: 270px; display: none">
								<asp:Table ID="HistoryTable" runat="server" Width="310">
								</asp:Table>
							</span>
						</td>
					</tr>
					<tr>
						<td class="whitecol" colspan="2" align="center">
							<asp:Button ID="btnSearch1" Style="z-index: 0" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
						</td>
					</tr>
				</table>
				<table id="ShowDataTable" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td align="center">
							<asp:DataGrid ID="DataGridC1" runat="server" PagerStyle-Visible="False" AutoGenerateColumns="False" AllowPaging="true" AllowSorting="true" CssClass="font" Width="100%">
								<AlternatingItemStyle BackColor="#EEEEEE" />
								<HeaderStyle CssClass="head_navy" />
								<Columns>
									<asp:BoundColumn HeaderText="序號"></asp:BoundColumn>
									<asp:BoundColumn DataField="Orgname" HeaderText="機構名稱"></asp:BoundColumn>
									<asp:BoundColumn DataField="classcname2" HeaderText="班名"></asp:BoundColumn>
									<asp:BoundColumn DataField="stdate" HeaderText="訓練起日"></asp:BoundColumn>
									<asp:BoundColumn DataField="ftdate" HeaderText="訓練迄日"></asp:BoundColumn>
									<asp:TemplateColumn HeaderText="功能">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center" Wrap="false" Font-Size="Small"></ItemStyle>
										<ItemTemplate>
											<asp:LinkButton ID="lbtAdd1" runat="server" Text="新增" CommandName="add1" CssClass="Linkbutton"></asp:LinkButton>&nbsp;
											<asp:LinkButton ID="lbtUpdate1" runat="server" Text="修改" CommandName="update1" CssClass="Linkbutton"></asp:LinkButton>&nbsp;
											<asp:LinkButton ID="lbtPrint1" runat="server" Text="列印" CommandName="print1" CssClass="asp_Export_M"></asp:LinkButton>&nbsp;
										</ItemTemplate>
									</asp:TemplateColumn>
								</Columns>
								<PagerStyle Visible="False"></PagerStyle>
							</asp:DataGrid>
							<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
						</td>
					</tr>
					<tr>
						<td align="center">
							<asp:DataGrid ID="DataGridS1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
								<AlternatingItemStyle BackColor="#F5F5F5"/>
								<HeaderStyle CssClass="head_navy" />
								<Columns>
									<asp:BoundColumn DataField="StudID" HeaderText="學號"></asp:BoundColumn>
									<asp:BoundColumn DataField="name" HeaderText="姓名"></asp:BoundColumn>
									<asp:BoundColumn DataField="Sex2" HeaderText="性別"></asp:BoundColumn>
									<asp:BoundColumn DataField="birthday" HeaderText="出生日期"></asp:BoundColumn>
									<asp:BoundColumn DataField="idno" HeaderText="身分證號碼"></asp:BoundColumn>
									<asp:BoundColumn DataField="BUDGETIDN" HeaderText="預算別"></asp:BoundColumn>
									<asp:BoundColumn DataField="vbcRatio" HeaderText="自行負擔費用比率"></asp:BoundColumn>
									<asp:BoundColumn DataField="defstdcost" HeaderText="應繳自行負擔費用"></asp:BoundColumn>
									<asp:TemplateColumn HeaderText="收據號碼">
										<ItemTemplate>
											<asp:TextBox ID="receipt" runat="server" MaxLength="100"></asp:TextBox>
											<asp:HiddenField ID="hidsocid" runat="server" />
										</ItemTemplate>
									</asp:TemplateColumn>
									<asp:BoundColumn DataField="MINAME" HeaderText="備註"></asp:BoundColumn>
								</Columns>
								<PagerStyle Visible="false"></PagerStyle>
							</asp:DataGrid>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td align="center">
				<asp:Label ID="labmsg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
			</td>
		</tr>
		<tr>
			<td>
				<p align="center">
					<asp:Button ID="btnSave1" runat="server" Text="存檔" CssClass="asp_button_S"></asp:Button>&nbsp;
					<asp:Button ID="btnBack1" runat="server" Text="回上頁" CssClass="asp_button_S"></asp:Button>&nbsp;
				</p>
			</td>
		</tr>
	</table>
	<asp:HiddenField ID="Hidocid" runat="server" />
	</form>
</body>
</html>
