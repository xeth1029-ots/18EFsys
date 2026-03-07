<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="EXAM_03_001_add.aspx.vb" Inherits="WDAIIP.EXAM_03_001_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>甄試班級維護</title>
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
		function choose_class() {
			openClass('../../SD/02/SD_02_ch.aspx?special=4&RID=' + document.form1.RIDValue.value);
		}

		function dg_create() {
			var msg = '';

			if (document.getElementById('ddl_cETID').value == '0' || document.getElementById('ddl_cETID').value == '') {
				msg += '請選擇【題組類別】內容\n';
			}

			if (document.getElementById('ddl_qtype').value == '0') {
				msg += '請選擇【題目類型】內容\n';
			}
			if ((document.getElementById('txt_num').value == '') || (document.getElementById('txt_num').value == '0')) {
				msg += '請填寫【題數】內容,【題數】不可為0\n';
			}
			else {
				if (!isUnsignedInt(document.getElementById('txt_num').value)) {
					msg += '【題數】填寫內容有誤\n';
				}
			}
			if ((document.getElementById('txt_score').value == '') || (document.getElementById('txt_score').value == '0')) {
				msg += '請填寫【總配分】,【總配分】不可為0內容\n';
			}
			else {
				if (!isUnsignedInt(document.getElementById('txt_score').value)) {
					msg += '【總配分】填寫內容有誤\n';
				}
			}
			if (msg != '') {
				window.alert(msg);
				return false;
			}
		}

		function check_asave() {
			var msg = '';
			var check_rbl = getValue("rbl_isonline");

			if (document.getElementById('OCID1').value == '') {
				msg = '請選擇【職類/班別】內容\n';
			}

			if (document.getElementById('txt_examtime').value == '') {
				msg += '請填選【考試時間】內容\n';
			}
			else {
				if (!isUnsignedInt(document.getElementById('txt_examtime').value)) {
					msg += '【考試時間】填寫內容有誤\n';
				}
			}

			if (check_rbl == 'Y') {
				if ((document.getElementById('txt_examdate1').value == '') || (document.getElementById('ddl_shour').value == '0') || (document.getElementById('ddl_ehour').value == '0') || (document.getElementById('ddl_sminute').value == '0') || (document.getElementById('ddl_eminute').value == '0')) {
					msg += '【線上登入時間】有選項未選擇\n';
				}
				else {
					if (document.getElementById('ddl_shour').value > document.getElementById('ddl_ehour').value) {
						msg += '【線上登入時間】起始時間不得大於結束時間\n';
					}
					if (document.getElementById('ddl_shour').value == document.getElementById('ddl_ehour').value) {
						if (document.getElementById('ddl_sminute').value > document.getElementById('ddl_eminute').value) {
							msg += '【線上登入時間】起始時間不得大於結束時間\n';
						}
					}
				}
			}

			if (msg != '') {
				window.alert(msg);
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
				<%--<table class="font" id="tab_title" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"><FONT face="新細明體">首頁&gt;&gt;招生甄試設定管理&gt;&gt;甄試班級考題設定&gt;&gt;</FONT></asp:Label><asp:Label ID="TitleLab2" runat="server"><font color="#990000">甄試班級維護</font></asp:Label>
						</td>
					</tr>
				</table>--%>
				<asp:Panel ID="tab_add" runat="server" Visible="False">
					<table class="table_sch" cellspacing="1" cellpadding="1">
						<tr>
							<td class="bluecol" width="20%">
								訓練機構
							</td>
							<td colspan="6" class="whitecol">
								<asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Button8" value="..." type="button" name="Button8" runat="server">
								<input id="RIDValue" type="hidden" name="RIDValue" runat="server"><br>
							</td>
						</tr>
						<tr>
							<td class="bluecol" width="20%">
								職類/班級
							</td>
							<td colspan="6" class="whitecol">
								<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
								<asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox><input id="Button5" value="..." type="button" name="Button5" runat="server">
								<input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
								<input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server"><br>
							</td>
						</tr>
						<tr>
							<td class="bluecol" width="20%">
								試卷類型
							</td>
							<td class="whitecol" colspan="6">
								<asp:RadioButtonList ID="rbl_isonline" runat="server" Width="20%" RepeatDirection="Horizontal" AutoPostBack="True" CellSpacing="0" CellPadding="0">
									<asp:ListItem Value="N" Selected="True">一般筆試</asp:ListItem>
									<asp:ListItem Value="Y">線上考試</asp:ListItem>
								</asp:RadioButtonList>
							</td>
						</tr>
						<tr>
							<td class="bluecol" width="20%">
								線上登入時間
							</td>
							<td class="whitecol" colspan="6">
								<asp:TextBox ID="txt_examdate1" runat="server" onfocus="this.blur()" Width="13%"></asp:TextBox>
								<asp:DropDownList ID="ddl_shour" runat="server" Enabled="False">
								</asp:DropDownList>
								：
								<asp:DropDownList ID="ddl_sminute" runat="server" Enabled="False">
								</asp:DropDownList>
								～
								<asp:DropDownList ID="ddl_ehour" runat="server" Enabled="False">
								</asp:DropDownList>
								：
								<asp:DropDownList ID="ddl_eminute" runat="server" Enabled="False">
								</asp:DropDownList>
								<font color="#ff0000">
									<br>
									(區間內未登入者將無法進行線上考試)</font></td>
						</tr>
						<tr>
							<td class="bluecol" width="20%">
								考試時間
							</td>
							<td class="whitecol" colspan="6">
								<asp:TextBox ID="txt_examtime" runat="server" Width="10%"></asp:TextBox>(時間單位:分)
							</td>
						</tr>
						<tr>
							<td class="bluecol" width="20%">
								出題順序
							</td>
							<td class="whitecol" colspan="6">
								<asp:RadioButtonList ID="rblSortType" Style="z-index: 0" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
									<asp:ListItem Value="1" Selected="True">依順序</asp:ListItem>
									<asp:ListItem Value="2">亂數</asp:ListItem>
								</asp:RadioButtonList>
							</td>
						</tr>
						<tr>
							<td class="bluecol" rowspan="2" width="20%">
								試卷內容
							</td>
							<td class="head_navy" align="center" width="15%">
								<strong><strong>題組類別</strong></strong>
							</td>
							<td class="head_navy" align="center" width="12.5%">
								<strong><strong>題組子類別</strong></strong>
							</td>
							<td class="head_navy" align="center" width="15%">
								<strong>題目類型</strong>
							</td>
							<td class="head_navy" align="center" width="12.5%">
								<strong>題數</strong>
							</td>
							<td class="head_navy" align="center" width="12.5%">
								<strong>總配分</strong>
							</td>
							<td class="head_navy" align="center" width="12.5%">
								<strong>功能</strong>
							</td>
						</tr>
						<tr>
							<td class="whitecol" align="center">
								<asp:DropDownList ID="ddl_pETID" runat="server" AutoPostBack="True"  Width="100%">
								</asp:DropDownList>
							</td>
							<td class="whitecol" align="center">
								<asp:DropDownList ID="ddl_cETID" runat="server"  Width="100%">
								</asp:DropDownList>
							</td>
							<td class="whitecol" align="center">
								<asp:DropDownList ID="ddl_qtype" runat="server"  Width="100%">
								</asp:DropDownList>
							</td>
							<td class="whitecol" align="center">
								<asp:TextBox ID="txt_num" runat="server" Width="100%"></asp:TextBox>
							</td>
							<td class="whitecol" align="center">
								<asp:TextBox ID="txt_score" runat="server" Width="100%"></asp:TextBox>
							</td>
							<td class="whitecol" align="center">
								<input id="hid_chkedit" type="hidden" name="hid_chkedit" runat="server">
                                <input id="hid_qtype" type="hidden" name="hid_qtype" runat="server">
                                <input id="hid_count" value="1" type="hidden" name="hid_count" runat="server">
								<asp:Button ID="btn_dgcrt" runat="server" Text="輸入" CssClass="asp_button_M"></asp:Button>
							</td>
						</tr>
					</table>
					<asp:DataGrid ID="dg_view" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="false" CellPadding="8">
						<AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
						<HeaderStyle CssClass="head_navy"></HeaderStyle>
						<Columns>
							<asp:BoundColumn DataField="id" HeaderText="順序">
								<HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn Visible="False" DataField="ecid" HeaderText="ecid"></asp:BoundColumn>
							<asp:BoundColumn Visible="False" DataField="etid" HeaderText="etid"></asp:BoundColumn>
							<asp:BoundColumn Visible="False" DataField="petid" HeaderText="petid"></asp:BoundColumn>
							<asp:BoundColumn DataField="pName" HeaderText="題組類別">
								<HeaderStyle HorizontalAlign="Center" Width="11%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="cName" HeaderText="題組子類別">
								<HeaderStyle HorizontalAlign="Center" Width="11%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn Visible="False" DataField="qtype" HeaderText="qtype"></asp:BoundColumn>
							<asp:BoundColumn DataField="qtype_name" HeaderText="題目類型">
								<HeaderStyle HorizontalAlign="Center" Width="11%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="num" HeaderText="題數">
								<HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="one_score" HeaderText="每題配分">
								<HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="score" HeaderText="總配分">
								<HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:BoundColumn DataField="total_num" HeaderText="題庫目">
								<HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
							</asp:BoundColumn>
							<asp:TemplateColumn HeaderText="功能">
								<HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
								<ItemStyle HorizontalAlign="Center"></ItemStyle>
								<ItemTemplate>
									<asp:Button ID="btn_aedit" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
									<asp:Button ID="btn_adel" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
								</ItemTemplate>
							</asp:TemplateColumn>
						</Columns>
						<PagerStyle Visible="False"></PagerStyle>
					</asp:DataGrid>
					<div align="right">
						&nbsp;
						<asp:Label ID="lbl_total" runat="server" Visible="False">合計總分：</asp:Label>
						<asp:Label ID="lbl_score" runat="server" Visible="False">0</asp:Label></div>
					<div align="center" class="whitecol">
						<asp:Button ID="btn_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button><font face="新細明體"> &nbsp;&nbsp;</font>
						<asp:Button ID="btn_lev" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button></div>
				</asp:Panel>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
