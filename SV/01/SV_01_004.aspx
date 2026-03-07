<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SV_01_004.aspx.vb" Inherits="WDAIIP.SV_01_004" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>問卷資料填寫</title>
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
		function GETvalue() {
			document.getElementById('Button7').click();
		}

		function choose_class() {
			openClass('../../SD/02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
		}


		 
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tr>
			<td>
				首頁&gt;&gt;系統管理&gt;&gt;問卷管理&gt;&gt;問卷資料填寫
			</td>
		</tr>
		<tr>
			<td>
				<table class="table_nw" id="table_Q" width="740" runat="server" cellspacing="1" cellpadding="1">
					<tr>
						<td class="bluecol" align="center" width="100">
							<label>
								<span class="style3">問卷名稱</span></label>
						</td>
						<td class="whitecol">
							<input id="Ipt_Name" style="width: 373px; height: 22px" maxlength="100" size="56" name="Ipt_Name" runat="server" height="18">
						</td>
					</tr>
					<tr align="center">
						<td colspan="2" class="whitecol">
							<%--<input id="search" type="button" value="查詢" name="search" runat="server" class="button_b_S" onclick="return search_onclick()">--%>
							<asp:Button ID="btnSearch1" runat="server" Text="查詢" CssClass="button_b_S" />
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table id="DataGrid1Table" width="740" runat="server">
					<tr width="100%">
						<td align="center" width="100%">
							<font face="新細明體">
								<asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" runat="server">
									<AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
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
												<asp:Button ID="Btn_edit" runat="server" Text="問卷填寫"></asp:Button>
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
			</td>
		</tr>
		<tr>
			<td>
				<table class="table_nw" id="Classtable" width="740" runat="server" cellpadding="1" cellspacing="1">
					<tr>
						<td class="bluecol" align="center" width="100">
							<label>
								<span class="style3">訓練機構</span></label>
						</td>
						<td class="whitecol">
							<asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="310px"></asp:TextBox>
							<input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
							<input id="RIDValue" type="hidden" name="Hidden2" runat="server">
							<asp:Button ID="Button7" Style="display: none" runat="server" Text="Button7"></asp:Button><br>
							<span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
								<asp:Table ID="HistoryRID" runat="server" Width="310px">
								</asp:Table>
							</span>
						</td>
					</tr>
					<tr class="SD_title">
						<td class="bluecol" align="center">
							職類/班別
						</td>
						<td class="whitecol">
							<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()"></asp:TextBox><asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()"></asp:TextBox>
							<input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
							<input id="TMIDValue1" style="width: 40px; height: 22px" type="hidden" name="Hidden1" runat="server">
							<input id="OCIDValue1" style="width: 32px; height: 22px" type="hidden" name="Hidden3" runat="server"><br>
							<span id="HistoryList" style="position: absolute; display: none; left: 270px">
								<asp:Table ID="HistoryTable" runat="server" Width="310">
								</asp:Table>
							</span>
						</td>
					</tr>
					<tr width="100%">
						<td align="center" colspan="2" class="whitecol">
							<asp:Button ID="btnSearch2" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table id="DataGrid2table" width="740" align="left" runat="server">
					<tr>
						<td align="center" width="100%">
							<asp:DataGrid ID="DataGrid2" runat="server" AllowPaging="True" AutoGenerateColumns="False" CssClass="font">
								<AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
								<HeaderStyle CssClass="head_navy"></HeaderStyle>
								<Columns>
									<asp:BoundColumn HeaderText="序號">
										<HeaderStyle Width="4%"></HeaderStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="orgname" HeaderText="訓練機構">
										<HeaderStyle Width="24%"></HeaderStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="CName" HeaderText="班級名稱">
										<HeaderStyle Width="24%"></HeaderStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="SEdate" HeaderText="開結訓日期">
										<HeaderStyle Width="18%"></HeaderStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="opencount" HeaderText="開訓人數">
										<HeaderStyle Width="4%"></HeaderStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="closecount" HeaderText="結訓人數">
										<HeaderStyle Width="4%"></HeaderStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="inputcount" HeaderText="填寫問卷人數">
										<HeaderStyle Width="6%"></HeaderStyle>
									</asp:BoundColumn>
									<asp:TemplateColumn HeaderText="功能">
										<HeaderStyle Width="6%"></HeaderStyle>
										<ItemTemplate>
											<asp:Button ID="EDIT" runat="server" Text="編輯"></asp:Button>
										</ItemTemplate>
									</asp:TemplateColumn>
									<%--<asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>--%>
								</Columns>
								<PagerStyle Visible="False"></PagerStyle>
							</asp:DataGrid><uc1:PageControler ID="Pagecontroler2" runat="server"></uc1:PageControler>
						</td>
					</tr>
					<tr>
						<td>
							<p align="center">
								<asp:Label ID="msg2" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
								<asp:Button ID="return1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button></p>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<table id="DataGrid3table" width="740" runat="server">
					<tr id="TR1" runat="server">
						<td style="height: 17px">
							<asp:Label ID="ORGL" runat="server" CssClass="font">訓練機構 :</asp:Label><asp:Label ID="ORGL2" runat="server" CssClass="font"></asp:Label>
						</td>
						<td style="height: 17px">
							<asp:Label ID="CLASSL" runat="server" CssClass="font">班級名稱 :</asp:Label><asp:Label ID="CLASSL2" runat="server" CssClass="font"></asp:Label>
						</td>
						<td style="height: 17px">
							<asp:Label ID="ODDATE" runat="server" CssClass="font">開結訓日期 :</asp:Label><asp:Label ID="ODDATE2" runat="server" CssClass="font"></asp:Label>
						</td>
					</tr>
					<tr id="TR2" runat="server">
						<td align="center" width="100%" colspan="3">
							<asp:DataGrid ID="DataGrid3" runat="server" Width="359px" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" AllowCustomPaging="True">
								<AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
								<HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
								<Columns>
									<asp:BoundColumn DataField="StudentID" HeaderText="學號">
										<HeaderStyle Width="15%"></HeaderStyle>
									</asp:BoundColumn>
									<asp:BoundColumn DataField="Sname" HeaderText="姓名">
										<HeaderStyle Width="40px"></HeaderStyle>
									</asp:BoundColumn>
									<asp:TemplateColumn HeaderText="功能">
										<HeaderStyle Width="45%"></HeaderStyle>
										<ItemTemplate>
											<asp:Button ID="InsertBtn" runat="server" Text="新增" CommandName="I"></asp:Button>
											<asp:Button ID="EditBtn" runat="server" Text="編輯" CommandName="E"></asp:Button>
											<asp:Button ID="DeleteBtn" runat="server" Text="刪除" CommandName="D"></asp:Button>
										</ItemTemplate>
									</asp:TemplateColumn>
									<asp:BoundColumn Visible="False" DataField="SOCID" HeaderText="SOCID"></asp:BoundColumn>
									<asp:BoundColumn Visible="False" DataField="isinput" HeaderText="是否有填問卷"></asp:BoundColumn>
								</Columns>
								<PagerStyle Visible="False"></PagerStyle>
							</asp:DataGrid><uc1:PageControler ID="Pagecontroler3" runat="server"></uc1:PageControler>
						</td>
					</tr>
					<tr align="center">
						<td colspan="3">
							<p align="center">
								<asp:Label ID="msg3" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
								<asp:Button ID="return2" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button><input id="SVID" type="hidden" name="SVID" runat="server"><input id="OCID2" type="hidden" name="SVID" runat="server"><input id="Sname" type="hidden" name="SVID" runat="server"><input id="StudentID" type="hidden" name="SVID" runat="server"></p>
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
