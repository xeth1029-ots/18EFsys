<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_022.aspx.vb" Inherits="WDAIIP.SD_05_022" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>SD_05_022</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script type="text/javascript">

		//個人訓練費用(含就業輔導費)
		function Getcost1(obj) {
			var Mytable = document.getElementById("DataGrid1");
			var MyLevel = Mytable.rows[1].cells[7].children[0];
			if (obj.checked) {
				if (MyLevel.value != '') {
					if (confirm('您要以第一位學員當預設值嗎?')) {
						for (var i = 2; i < Mytable.rows.length; i++) {
						    var MyText = Mytable.rows[i].cells[7].children[0];
							//if (!MyText.disabled){
							MyText.value = MyLevel.value;
							//}
						}
					}
					else {
						obj.checked = false;
					}
				}
				else {
					alert('未設定第一位學員的【個人訓練費用(含就業輔導費)】');
					obj.checked = false;
				}
			}
		}

		//個人訓練費用(不含就業輔導費)
		function Getcost2(obj) {
		    var Mytable = document.getElementById("DataGrid1");
		    var MyLevel = Mytable.rows[1].cells[8].children[0];
			if (obj.checked) {
				if (MyLevel.value != '') {
					if (confirm('您要以第一位學員當預設值嗎?')) {
						for (var i = 2; i < Mytable.rows.length; i++) {
						    var MyText = Mytable.rows[i].cells[8].children[0];
							//if (!MyText.disabled){
							MyText.value = MyLevel.value;
							//}
						}
					}
					else {
						obj.checked = false;
					}
				}
				else {
					alert('未設定第一位學員的【個人訓練費用(不含就業輔導費)】');
					obj.checked = false;
				}
			}
		}

		//第X次撥款數
		function Getcount(obj) {
			var Mytable = document.getElementById("DataGrid1");
			var MyLevel = Mytable.rows[1].cells[9].children[0];
			if (obj.checked) {
				if (MyLevel.value != '') {
					if (confirm('您要以第一位學員當預設值嗎?')) {
						for (var i = 2; i < Mytable.rows.length; i++) {
						    var MyText = Mytable.rows[i].cells[9].children[0];
							//if (!MyText.disabled){
							MyText.value = MyLevel.value;
							//}
						}
					}
					else {
						obj.checked = false;
					}
				}
				else {
					alert('未設定第一位學員的【撥款數】');
					obj.checked = false;
				}
			}
		}

		//備註：受訓狀況(中長期失業週數)
		function Getnote(obj) {
			var Mytable = document.getElementById("DataGrid1");
			var MyLevel = Mytable.rows[1].cells[10].children[0];
			if (obj.checked) {
				if (MyLevel.value != '') {
					if (confirm('您要以第一位學員當預設值嗎?')) {
						for (var i = 2; i < Mytable.rows.length; i++) {
						    var MyText = Mytable.rows[i].cells[10].children[0];
							//if (!MyText.disabled){
							MyText.value = MyLevel.value;
							//}
						}
					}
					else {
						obj.checked = false;
					}
				}
				else {
					alert('未設定第一位學員的【備註：受訓狀況(中長期失業週數)】');
					obj.checked = false;
				}
			}
		}


		function GETvalue() {
			document.getElementById('Button3').click();
		}

		function choose_class() {
			openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
		}

		function CheckData() {
			var msg = '';
			if (document.form1.OCIDValue1.value == '') msg += '請選擇職類班別\n';

			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		function IsDate(MyDate) {
			if (MyDate != '') {
				if (!checkDate(MyDate))
					return false;
			}
			return true;
		}

		function open_History(idnoValue) {
			//debugger;
			//window.open('../01/SD_01_001_old.aspx?state=2&IDNO='+idnoValue+httpValue,'history','width=700,height=500,scrollbars=1')
			window.open('../01/SD_01_001_old.aspx?IDNO=' + idnoValue, 'history', 'width=700,height=500,scrollbars=1')
		}

		function SelectAll() {
			var MyTable = document.getElementById('DataGrid1');

			for (i = 1; i < MyTable.rows.length; i++) {
				if (!MyTable.rows[i].cells[0].children[0].disabled)
					MyTable.rows[i].cells[0].children[0].checked = MyTable.rows[0].cells[0].children[0].checked;
			}
		}

		function SET_Rate(obj) {
			//debugger;
			document.form1.hidRate.value = obj.value;
		}

	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="740" border="0">
		<tr>
			<td>
				<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										<FONT face="新細明體">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">個人訓練費用明細表</font></FONT>
							</asp:Label>
						</td>
					</tr>
				</table>
				<asp:Panel ID="TableSearch" runat="server">
					<table class="table_sch" cellpadding="1" cellspacing="1">
						<tr>
							<td width="100" class="bluecol">
								訓練機構
							</td>
							<td class="whitecol">
								<asp:TextBox ID="center" runat="server" AutoPostBack="True" Width="410px"></asp:TextBox><input id="RIDValue" type="hidden" name="RIDValue" runat="server">
								<input id="BtnOrg" type="button" value="..." name="BtnOrg" runat="server"><br>
								<asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
								<span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
									<asp:Table ID="HistoryRID" runat="server" Width="310px">
									</asp:Table>
								</span>
							</td>
						</tr>
						<tr>
							<td width="100" class="bluecol_need">
								職類/班別<font color="#ff0000">*</font>
							</td>
							<td class="whitecol">
								<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox><asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox><input onclick="choose_class()" type="button" value="...">
								<input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
								<input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
								<font color="#ffffff"></font>
								<br>
								<span id="HistoryList" style="position: absolute; display: none; left: 270px">
									<asp:Table ID="HistoryTable" runat="server" Width="310">
									</asp:Table>
								</span>
							</td>
						</tr>
					</table>
					<table width="100%">
						<tr>
							<td class="whitecol">
								<p align="center">
									<asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="btn_add" runat="server" Text="新增" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="print_btn" runat="server" Text="列印" CommandName="print" CssClass="asp_Export_M"></asp:Button></p>
							</td>
						</tr>
						<tr>
							<td align="center" class="whitecol">
								<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
							</td>
						</tr>
					</table>
				</asp:Panel>
			</td>
		</tr>
	</table>
	<table class="font" id="TableShowData" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
		<tr>
			<td align="right">
				<font class="font">已使用撥款百分比比率:</font><asp:Label ID="labRateAdd" runat="server" ForeColor="Red"></asp:Label>
			</td>
		</tr>
		<tr>
			<td>
				<asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
					<AlternatingItemStyle BackColor="#F5F5F5" />
					<HeaderStyle CssClass="head_navy" />
					<Columns>
						<asp:TemplateColumn HeaderText="選取">
							<HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<HeaderTemplate>
								<input onclick="SelectAll();" type="checkbox">
							</HeaderTemplate>
							<ItemTemplate>
								<input id="Checkbox1" type="checkbox" name="Checkbox1" runat="server">
								<input id="IDNO" type="hidden" runat="server" value='<%# DataBinder.Eval(Container, "DataItem.IDNO") %>'>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:BoundColumn DataField="StudentID" HeaderText="學號">
							<HeaderStyle HorizontalAlign="Center" Width="40px"></HeaderStyle>
							<ItemStyle Wrap="False" HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:TemplateColumn HeaderText="姓名">
							<ItemStyle Wrap="False" HorizontalAlign="Center" Width="60px"></ItemStyle>
							<ItemTemplate>
								<asp:Label ID="labName" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.Name") %>'>
								</asp:Label>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
							<HeaderStyle Width="90px"></HeaderStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Sex" HeaderText="性別">
							<HeaderStyle Width="40px"></HeaderStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Birthday" HeaderText="出生日期" DataFormatString="{0:d}">
							<HeaderStyle Width="80px"></HeaderStyle>
						</asp:BoundColumn>
						<asp:TemplateColumn HeaderText="學員狀態">
							<HeaderStyle Wrap="False"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center" Width="60px"></ItemStyle>
							<HeaderTemplate>
								學員狀態<br>
								(前次撥款費用)
							</HeaderTemplate>
							<ItemTemplate>
								<asp:Label ID="labStudStatus" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.StudStatus") %>'>
								</asp:Label>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:TemplateColumn HeaderText="個人訓練費用">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<HeaderTemplate>
								個人訓練費用<br>
								(含就業輔導費)<input id="cost1" type="checkbox" onclick="Getcost1(this);">
							</HeaderTemplate>
							<ItemTemplate>
								<asp:TextBox ID="txtJobCost" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.JobCost") %>' Width="90px" CssClass="center1">
								</asp:TextBox>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:TemplateColumn HeaderText="個人訓練費用">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<HeaderTemplate>
								個人訓練費用<br>
								(不含就業輔導費)<input id="cost2" type="checkbox" onclick="Getcost2(this);">
							</HeaderTemplate>
							<ItemTemplate>
								<asp:TextBox ID="txtOtherJobCost" runat="server" Width="90px" Text='<%# DataBinder.Eval(Container, "DataItem.OtherJobCost") %>' CssClass="center1">
								</asp:TextBox>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:TemplateColumn HeaderText="第X次撥款數">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<HeaderTemplate>
								第
								<asp:Label ID="labTimes" runat="server">X</asp:Label>次撥款數<font color="#ff0000">*</font><br />
								(
								<asp:TextBox ID="txtRate" runat="server" Width="30px" MaxLength="3" CssClass="center1"></asp:TextBox>%)<input id="count" type="checkbox" onclick="Getcount(this);">
							</HeaderTemplate>
							<ItemTemplate>
								<asp:TextBox ID="txtCost" runat="server" Width="90px" Text='<%# DataBinder.Eval(Container, "DataItem.Cost") %>' CssClass="center1">
								</asp:TextBox>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:TemplateColumn HeaderText="備註：受訓狀況">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<HeaderTemplate>
								備註：受訓狀況<br>
								(中長期失業週數)<input id="note" type="checkbox" onclick="Getnote(this);">
							</HeaderTemplate>
							<ItemTemplate>
								<asp:TextBox ID="txtState" runat="server" Width="90px" MaxLength="50" Text='<%# DataBinder.Eval(Container, "DataItem.State") %>' CssClass="center1">
								</asp:TextBox>
							</ItemTemplate>
						</asp:TemplateColumn>
					</Columns>
				</asp:DataGrid>
			</td>
		</tr>
		<tr>
			<td align="center">
				<asp:Button ID="btn_Save2" runat="server" Text="儲存" CssClass="asp_button_S"></asp:Button>&nbsp;<asp:Button ID="btn_back" runat="server" Text="回上一頁" CssClass="asp_button_S"></asp:Button><input id="hidRate" type="hidden" name="hidRate" runat="server"><input id="times2" type="hidden" name="times2" runat="server">
				<%--<uc1:pagecontroler id="PageControler1" runat="server"></uc1:pagecontroler>--%>
			</td>
		</tr>
	</table>
	<p>
		&nbsp;
		<%--<TABLE class="font" id="DataGridTable2" cellSpacing="1" cellPadding="1" width="100%" border="0"
					runat="server">
					<tr>
						<td>--%>
		<asp:DataGrid ID="DataGrid2" runat="server" Width="740px" AutoGenerateColumns="False" Height="29px">
			<SelectedItemStyle Font-Size="Smaller"></SelectedItemStyle>
			<EditItemStyle Font-Size="Smaller"></EditItemStyle>
			<AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
			<HeaderStyle CssClass="head_navy"></HeaderStyle>
			<Columns>
				<asp:BoundColumn DataField="Classname" HeaderText="班級名稱"></asp:BoundColumn>
				<asp:BoundColumn DataField="Times" HeaderText="撥款次數"></asp:BoundColumn>
				<asp:BoundColumn DataField="Rate" HeaderText="撥款比率"></asp:BoundColumn>
				<asp:TemplateColumn HeaderText="功能">
					<ItemTemplate>
						<asp:Button ID="edit_btn" runat="server" Text="編輯" CommandName="edit"></asp:Button>
					</ItemTemplate>
				</asp:TemplateColumn>
			</Columns>
		</asp:DataGrid>
		<%--<td></td>
					</tr>
				</TABLE>--%>
	</p>
	<p>
		<font face="新細明體"></font>&nbsp;</p>
	</form>
</body>
</html>
