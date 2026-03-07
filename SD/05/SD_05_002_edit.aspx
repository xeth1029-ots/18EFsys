<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_002_edit.aspx.vb" Inherits="WDAIIP.SD_05_002_edit" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>學員出缺勤作業</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
	<script type="text/javascript" language="javascript">
		function w_printS1() {
			window.print();
		}

		function check_data() {
			var mytable = document.getElementById('DataGrid1');
			var msg = '';
			for (var i = 1; i < mytable.rows.length; i++) {
			    var mydate = mytable.rows[i].cells[1].children[0]; //mydrop = 第i行的[申請日期]那個欄位
				var myLeaveID = mytable.rows[i].cells[2].children[0]; //mydrop = 第i行的[假別]那個欄位

				if (mydate.value == '') msg += '請輸入申請日期(第' + i + '行)\n';
				else if (!checkDate(mydate.value)) msg += '請申請日期不符合日期格式(第' + i + '行)\n';
				if (myLeaveID.selectedIndex == 0) msg += '請選擇假別(第' + i + '行)\n';
			}

			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		//function btnPrint1_onclick() {		}
		//w_printS1

	</script>
</head>
<body onload="FrameLoad();">
	<form id="form1" method="post" runat="server">
	<table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
		<tr>
			<td>
				<%--<div id="block">--%>
					<table id="EditTable" cellspacing="1" cellpadding="1" width="100%" border="0">
						<tr>
							<td>
								<table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
									<tr>
										<td class="bluecol" width="20%">訓練機構 </td>
										<td class="whitecol" width="30%">
											<asp:Label ID="OrgName" runat="server"></asp:Label>
										</td>
										<td class="bluecol" width="20%">班別 </td>
										<td class="whitecol" width="30%">
											<asp:Label ID="ClassCName" runat="server"></asp:Label>
										</td>
									</tr>
									<tr>
										<td class="bluecol">學員姓名 </td>
										<td class="whitecol">
											<asp:Label ID="Name" runat="server"></asp:Label>
										</td>
										<td class="bluecol">學號 </td>
										<td class="whitecol">
											<asp:Label ID="StudentID" runat="server"></asp:Label>                                            
										</td>
									</tr>
									<tr>
										<td class="bluecol">學員狀態 </td>
										<td colspan="3" class="whitecol">
											<asp:Label ID="StudStatus" runat="server"></asp:Label>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td>
								<asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
									<%--<FooterStyle BackColor="#E7FFE7"></FooterStyle>--%>
									<AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>									
									<HeaderStyle CssClass="head_navy"></HeaderStyle>
									<Columns>
										<asp:TemplateColumn HeaderText="功能">
											<HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center" ></ItemStyle>
											<ItemTemplate>
												<asp:Button ID="Button1" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
												<input id="hid_leaveT" runat="server" type="hidden">
												<asp:HiddenField ID="Hid_SeqNo" runat="server" />
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="點名日期">
                                            <HeaderStyle HorizontalAlign="Center" ></HeaderStyle>
                                            <ItemStyle CssClass="whitecol" Wrap="false"/>
											<ItemTemplate>
												<asp:TextBox ID="LeaveDate" runat="server" Width="77%" MaxLength="11"></asp:TextBox>
                                                <img id="Img1" style="cursor: pointer" onclick="" alt="" src="../../images/show-calendar.gif" align="top" runat="server" width="30" height="30">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="假別">
											<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
											<ItemTemplate>
												<asp:DropDownList ID="LeaveID" runat="server" AutoPostBack="True">
												</asp:DropDownList>
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節1">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C1" type="checkbox" name="C1" runat="server">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節2">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C2" type="checkbox" runat="server" name="C2">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節3">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C3" type="checkbox" runat="server" name="C3">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節4">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C4" type="checkbox" runat="server" name="C4">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節5">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C5" type="checkbox" runat="server" name="C5">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節6">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C6" type="checkbox" runat="server" name="C6">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節7">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C7" type="checkbox" runat="server" name="C7">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節8">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C8" type="checkbox" runat="server" name="C8">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節9">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C9" type="checkbox" runat="server" name="C9">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節10">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C10" type="checkbox" runat="server" name="C10">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節11">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C11" type="checkbox" name="C11" runat="server">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:TemplateColumn HeaderText="節12">
											<HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="C12" type="checkbox" runat="server" name="C12">
											</ItemTemplate>
										</asp:TemplateColumn>
										<asp:BoundColumn DataField="Hours" HeaderText="總計時數">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
										</asp:BoundColumn>
										<%--<asp:BoundColumn Visible="False" DataField="SeqNo" HeaderText="SeqNo"></asp:BoundColumn>--%>
										<asp:TemplateColumn HeaderText="不列入&lt;br/&gt;缺曠課">
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
											<ItemStyle HorizontalAlign="Center"></ItemStyle>
											<ItemTemplate>
												<input id="TurnoutIgnore" type="checkbox" runat="server">
											</ItemTemplate>
										</asp:TemplateColumn>
									</Columns>
								</asp:DataGrid>
							</td>
						</tr>
					</table>
				<%--</div>--%>
				<table id="EditTable2" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td align="center" class="whitecol">
							<asp:Button ID="Button5" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
							<input id="Button6" type="button" value="回上一頁" runat="server" class="asp_button_M">
							<input id="btnPrint1" type="button" value="列印" runat="server" class="asp_Export_M">
						</td>
					</tr>
				</table>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
