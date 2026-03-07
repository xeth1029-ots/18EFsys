<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_09_010_R.aspx.vb" Inherits="WDAIIP.SD_09_010_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>SD_09_010_R</title>
	<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
	<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
	<meta name="vs_defaultClientScript" content="JavaScript" />
	<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
	<link href="../../css/style.css" type="text/css" rel="stylesheet" />
	<script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
	<script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
	<script type="text/javascript" language="javascript">
		function GETvalue() {
			document.getElementById('Button3').click();
		}
		function choose_class() {
			var RIDValue = document.getElementById('RIDValue');
			openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
		}

		function search1() {
			var msg = '';
			var OCIDValue1 = document.getElementById('OCIDValue1');
			if (OCIDValue1.value == '') msg += '請選擇職類班別\n';
			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		function ReportPrint(url) {
			var MyTable = document.getElementById('DG_stud');
			var OCIDValue1 = document.getElementById('OCIDValue1');
			var flag = false;
			var idno = '';
			for (i = 1; i < MyTable.rows.length; i++) {
				var MyCheck = MyTable.rows[i].cells[0].children[0];
				if (MyCheck.checked) {
					flag = true;
					if (idno != '') { idno += ','; }
					idno = '\'' + MyCheck.value + '\'';
				}
			}
			if (flag) {
				window.open('../../SQControl.aspx?' + url + '&IDNO=' + idno + '&OCID=' + OCIDValue1.value, 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1')
			}
			else
				alert('請先勾選學員。');

			return false;
		}			

	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;外籍人士上課時數紀錄</asp:Label>
                </td>
            </tr>
        </table>
	<table id="Frame" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
		<tbody>
			<tr>
				<td align="center">
					<%--<table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
						<tr>
							<td>
								<p>
									首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;<font color="#990000">外籍人士上課時數紀錄</font>&nbsp;</p>
							</td>
						</tr>
					</table>--%>
					<table id="Table3" class="table_sch" cellspacing="1" cellpadding="1">
						<tr>
							<td class="bluecol" style="width:20%">
								訓練機構
							</td>
							<td class="whitecol">
								<asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
								<input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
								<input type="button" value="..." id="Button2" name="Button2" runat="server" class="button_b_Mini" /><br />
								<asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
								<span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
									<asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
								</span>
							</td>
						</tr>
						<tr>
							<td class="bluecol_need">
								班別/職類
							</td>
							<td class="whitecol">
								<asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
								<asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
								<input type="button" value="..." onclick="choose_class()" class="button_b_Mini" />
								<input id="OCIDValue1" type="hidden" name="Hidden2" runat="server" />
								<input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" /><br />
								<span id="HistoryList" style="z-index: 101; position: absolute; display: none; left: 270px">
									<asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
								</span>
							</td>
						</tr>
					</table>
					<p align="center" class="whitecol">
						<asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
					</p>
					<table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server" class="font">
						<tr>
							<td>
								<asp:DataGrid ID="DG_stud" runat="server" Width="100%" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" CssClass="font" AllowCustomPaging="True" CellPadding="8">
									<AlternatingItemStyle BackColor="#F5F5F5" />
									<HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
									<Columns>
										<asp:TemplateColumn HeaderText="選取">
                                            <HeaderStyle Width="8%"/>
                                            <ItemStyle HorizontalAlign="Center"/>
											<ItemTemplate>
												<input id="Checkbox1" type="checkbox" runat="server" name="Checkbox1" />
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:BoundColumn DataField="StudentID" HeaderText="學號" HeaderStyle-Width="23%"></asp:BoundColumn>
										<asp:BoundColumn DataField="Name" HeaderText="姓名" HeaderStyle-Width="23%"></asp:BoundColumn>
										<asp:BoundColumn DataField="EngName" HeaderText="英文姓名" HeaderStyle-Width="23%"></asp:BoundColumn>
										<asp:BoundColumn DataField="StudStatus" HeaderText="學員狀態" HeaderStyle-Width="23%"></asp:BoundColumn>
										<asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
										<asp:BoundColumn Visible="False" DataField="StudentID" HeaderText="StudentID"></asp:BoundColumn>
									</Columns>
									<PagerStyle Visible="False"></PagerStyle>
								</asp:DataGrid>
							</td>
						</tr>
						<tr>
							<td align="center" class="whitecol">
								<asp:Button ID="submit" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button>
							</td>
						</tr>
					</table>
					<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
				</td>
			</tr>
		</tbody>
	</table>
	</form>
</body>
</html>
