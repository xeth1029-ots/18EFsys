<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_031.aspx.vb" Inherits="WDAIIP.SD_05_031" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title>學習護照</title>
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript" src="../../js/common.js"></script>
	<script src="../../js/common.js"></script>
	<script type="text/javascript">
	</script>
    <style type="text/css">
        .auto-style1 {
            color: Black;
            text-align: right;
            padding: 4px 6px;
            background-color: #f1f9fc;
            border-right: 3px solid #49cbef;
            height: 29px;
        }
        .auto-style2 {
            color: #333333;
            padding: 4px;
            height: 29px;
        }
    </style>
</head>
<body>
	<form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;學習護照</asp:Label>
                </td>
            </tr>
        </table>
	<table class="font" id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
		<%--<tr>
			<td>
				<table class="font" id="table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">學習護照</font>
							</asp:Label>
						</td>
					</tr>
				</table>
			</td>
		</tr>--%>
		<tr>
			<td>
				<table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
					<tr>
						<td class="bluecol" style="width:20%">
							身分證號碼
						</td>
						<td class="whitecol" style="width:30%">
							<asp:TextBox ID="sIDNO" runat="server"></asp:TextBox>
						</td>
						<td class="bluecol" style="width:20%">
							姓名
						</td>
						<td class="whitecol" style="width:30%">
							<asp:TextBox ID="sName" runat="server" Width="40%"></asp:TextBox>
						</td>
					</tr>
					<tr>
						<td class="bluecol">
							出生日期
						</td>
						<td class="whitecol" colspan="3">
							<font color="#ffffff">
								<asp:TextBox ID="sbirthday" runat="server" Columns="10" MaxLength="12" Width="15%"></asp:TextBox>
								<img style="cursor: pointer" onclick="javascript:show_calendar('sbirthday','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
							<%--<asp:label id="Note" runat="server" ForeColor="Red">搜尋條件【身分證號碼】與【姓名】，請擇一輸入</asp:label>--%>
						</td>
					</tr>
					<tr>
						<td class="whitecol" colspan="4" align="center">
							&nbsp;<asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_M" />
						</td>
					</tr>
					<tr>
						<td class="whitecol" colspan="4" align="center">
							<asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
						</td>
					</tr>
				</table>
			</td>
		</tr>
		<tr>
			<td>
				<asp:Panel ID="Panelshow1" runat="server" Visible="False">
					<table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
						<tr>
							<td colspan="4" class="head_navy">
								基本資料
							</td>
						</tr>
						<tr>
							<td class="bluecol" width="20%">
								姓名
							</td>
							<td class="whitecol">
								<asp:Label ID="labname" runat="server"></asp:Label>
								&nbsp;
							</td>
							<td class="bluecol" width="20%">
								身分證號碼
							</td>
							<td class="whitecol">
								<asp:Label ID="labidno" runat="server"></asp:Label>
								&nbsp;
							</td>
						</tr>
						<tr>
							<td class="bluecol" width="20%">
								生日
							</td>
							<td class="whitecol">
								<asp:Label ID="labbirthday" runat="server"></asp:Label>
								&nbsp;
							</td>
							<td class="bluecol">
								聯絡電話
							</td>
							<td class="whitecol">
								<asp:Label ID="labtel" runat="server"></asp:Label>
								&nbsp;
							</td>
						</tr>
						<tr>
							<td class="auto-style1">
								地址
							</td>
							<td class="auto-style2" colspan="3">
								<asp:Label ID="labaddress" runat="server"></asp:Label>
								&nbsp;
							</td>
						</tr>
					</table>
					<table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
						<tr>
							<td class="head_navy">
								參訓歷程&nbsp;
							</td>
						</tr>
						<tr>
							<td class="whitecol">
								<asp:Label ID="labmsg1" runat="server" ForeColor="Red"></asp:Label>
								&nbsp;<asp:DataGrid ID="DataGrid1" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" CssClass="font" PageSize="20" Width="100%" CellPadding="8">
									<AlternatingItemStyle BackColor="#F5F5F5" />
									<HeaderStyle CssClass="head_navy" />
									<Columns>
										<%--<asp:BoundColumn DataField="DistName" HeaderText="轄區中心" SortExpression="DistName">--%>
                                        <asp:BoundColumn DataField="DistName" HeaderText="轄區分署" SortExpression="DistName">
											<HeaderStyle ForeColor="#00ffff" HorizontalAlign="Center" Width="10%" />
											<ItemStyle HorizontalAlign="Center" />
										</asp:BoundColumn>
										<asp:BoundColumn DataField="Years" HeaderText="年度">
											<HeaderStyle Width="5%" />
										</asp:BoundColumn>
										<asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
											<HeaderStyle HorizontalAlign="Center" Width="10%" />
										</asp:BoundColumn>
										<asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" SortExpression="OrgName">
											<HeaderStyle ForeColor="#00ffff" HorizontalAlign="Center" Width="10%" />
										</asp:BoundColumn>
										<asp:BoundColumn DataField="TMID" HeaderText="訓練職類" SortExpression="TMID">
											<HeaderStyle ForeColor="#00ffff" HorizontalAlign="Center" Width="10%" />
										</asp:BoundColumn>
										<asp:BoundColumn DataField="CJOB_NAME" HeaderText="通俗職類">
											<HeaderStyle HorizontalAlign="Center" Width="10%" />
										</asp:BoundColumn>
										<asp:BoundColumn DataField="ClassName" HeaderText="班別名稱" SortExpression="ClassName">
											<HeaderStyle ForeColor="#00ffff" HorizontalAlign="Center" Width="10%" />
										</asp:BoundColumn>
										<asp:BoundColumn DataField="THours" HeaderText="受訓&lt;BR&gt;時數">
											<HeaderStyle HorizontalAlign="Center" Width="5%" />
										</asp:BoundColumn>
										<asp:BoundColumn DataField="TRound" HeaderText="受訓期間" SortExpression="TRound">
											<HeaderStyle ForeColor="#00ffff" HorizontalAlign="Center" Width="10%" />
											<ItemStyle HorizontalAlign="Center" />
										</asp:BoundColumn>
										<asp:BoundColumn DataField="WEEKS" HeaderText="上課時間">
											<HeaderStyle HorizontalAlign="Center" Width="15%"/>
										</asp:BoundColumn>
										<asp:BoundColumn DataField="TFlag" HeaderText="訓練&lt;BR&gt;狀態">
											<HeaderStyle HorizontalAlign="Center" Width="5%" />
										</asp:BoundColumn>
									</Columns>
									<PagerStyle Visible="False" />
								</asp:DataGrid>
							</td>
						</tr>
					</table>
					<table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
						<tr>
							<td class="head_navy">
								津貼歷程
							</td>
						</tr>
						<tr>
							<td class="whitecol">								
								<asp:Label ID="labmsg2" runat="server" ForeColor="Red"></asp:Label>
								<asp:DataGrid ID="DataGrid2" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" CssClass="font" PageSize="20" Width="100%" CellPadding="8">
									<AlternatingItemStyle BackColor="#F5F5F5" />
									<HeaderStyle CssClass="head_navy" />
									<Columns>
										<asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="30%"></HeaderStyle>
										</asp:BoundColumn>
										<asp:BoundColumn DataField="ClassCName" HeaderText="參訓課程">
                                            <HeaderStyle Width="30%"></HeaderStyle>
										</asp:BoundColumn>
										<asp:BoundColumn DataField="STDate" HeaderText="開訓日期" DataFormatString="{0:d}">
											<HeaderStyle Width="10%"></HeaderStyle>
										</asp:BoundColumn>
										<asp:BoundColumn DataField="FTDate" HeaderText="結訓日期" DataFormatString="{0:d}">
											<HeaderStyle Width="10%"></HeaderStyle>
										</asp:BoundColumn>
										<asp:BoundColumn DataField="TrainingMoney" HeaderText="申請補助金額">
											<HeaderStyle Width="10%"></HeaderStyle>
										</asp:BoundColumn>
										<asp:BoundColumn DataField="PayMoney" HeaderText="實領核發金額">
                                            <HeaderStyle Width="10%"></HeaderStyle>
										</asp:BoundColumn>
									</Columns>
									<PagerStyle Visible="False" />
								</asp:DataGrid>
							</td>
						</tr>
					</table>
					<table class="table_nw" cellspacing="1" cellpadding="1" width="740">
						<tr>
							<td class="head_navy">
								就業服務&nbsp;
							</td>
						</tr>
						<tr>
							<td class="whitecol">
								&nbsp;
								<asp:Label ID="labmsg3" runat="server" ForeColor="Red"></asp:Label>
								<asp:DataGrid ID="DataGrid3" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" CssClass="font" PageSize="20" Width="100%" CellPadding="8">
									<AlternatingItemStyle BackColor="#F5F5F5" />
									<HeaderStyle CssClass="head_navy" />
									<Columns>
										<asp:BoundColumn DataField="seqno" HeaderText="項次"></asp:BoundColumn>
										<asp:BoundColumn DataField="interDate1" HeaderText="介紹日期"></asp:BoundColumn>
										<asp:BoundColumn DataField="orgName" HeaderText="求才單位"></asp:BoundColumn>
										<asp:BoundColumn DataField="workName" HeaderText="職稱名稱"></asp:BoundColumn>
										<asp:BoundColumn DataField="workYN" HeaderText="僱用與否"></asp:BoundColumn>
										<asp:BoundColumn DataField="NRReason" HeaderText="僱主未能錄用原因"></asp:BoundColumn>
										<asp:BoundColumn DataField="NWReason" HeaderText="求職未能推介原因"></asp:BoundColumn>
									</Columns>
									<PagerStyle Visible="False" />
								</asp:DataGrid>
							</td>
						</tr>
					</table>
					<table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
						<tr>
							<td class="head_navy">
								技能檢定
							</td>
						</tr>
						<tr>
							<td class="whitecol">
								<asp:Label ID="Labmsg4" runat="server" ForeColor="Red"></asp:Label>
								<asp:DataGrid ID="DataGrid4" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" CssClass="font" PageSize="20" Width="100%" CellPadding="8">
									<AlternatingItemStyle BackColor="#F5F5F5" />
									<HeaderStyle CssClass="head_navy" />
									<Columns>
										<asp:BoundColumn DataField="seqno" HeaderText="項次" HeaderStyle-Width="25%"></asp:BoundColumn>
										<asp:BoundColumn DataField="name" HeaderText="職類名稱" HeaderStyle-Width="25%"></asp:BoundColumn>
										<asp:BoundColumn DataField="class" HeaderText="級別" HeaderStyle-Width="25%"></asp:BoundColumn>
										<asp:BoundColumn DataField="appday" HeaderText="發證日" HeaderStyle-Width="25%"></asp:BoundColumn>
									</Columns>
									<PagerStyle Visible="False" />
								</asp:DataGrid>
							</td>
						</tr>
					</table>
					<table class="table_nw" cellspacing="1" cellpadding="1" width="740">
						<tr>
							<td class="head_navy">
								工作經歷&nbsp;
							</td>
						</tr>
						<tr>
							<td class="whitecol">
								&nbsp;<asp:Label ID="Labmsg5" runat="server" ForeColor="Red"></asp:Label>
								<asp:DataGrid ID="DataGrid5" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" CssClass="font" PageSize="20" Width="100%" CellPadding="8">
									<AlternatingItemStyle BackColor="#F5F5F5" />
									<HeaderStyle CssClass="head_navy" />
									<Columns>
										<asp:BoundColumn DataField="seqno" HeaderText="項次" HeaderStyle-Width="5%"></asp:BoundColumn>
										<asp:BoundColumn DataField="OrgName" HeaderText="公司機稱" HeaderStyle-Width="15%"></asp:BoundColumn>
										<asp:BoundColumn DataField="trainName" HeaderText="行業別" HeaderStyle-Width="15%"></asp:BoundColumn>
										<asp:BoundColumn DataField="workName" HeaderText="工作職稱" HeaderStyle-Width="15%"></asp:BoundColumn>
										<asp:BoundColumn DataField="workName2" HeaderText="工作職稱" HeaderStyle-Width="15%"></asp:BoundColumn>
										<asp:BoundColumn DataField="workMemo" HeaderText="工作說明" HeaderStyle-Width="15%"></asp:BoundColumn>
										<asp:BoundColumn DataField="salary" HeaderText="薪資" HeaderStyle-Width="5%"></asp:BoundColumn>
										<asp:BoundColumn DataField="workTime" HeaderText="工作時間起迄" HeaderStyle-Width="15%"></asp:BoundColumn>
									</Columns>
									<PagerStyle Visible="False" />
								</asp:DataGrid>
							</td>
						</tr>
					</table>
					<table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
						<tr>
							<td class="head_navy">
								推介歷程
							</td>
						</tr>
						<tr>
							<td class="whitecol">
								<asp:Label ID="Labmsg6" runat="server" ForeColor="Red"></asp:Label>
								<asp:DataGrid ID="DataGrid6" runat="server" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False" CssClass="font" PageSize="20" Width="100%" CellPadding="8">
									<AlternatingItemStyle BackColor="#F5F5F5" />
									<HeaderStyle CssClass="head_navy" />
									<Columns>
										<asp:BoundColumn DataField="seqno" HeaderText="項次" HeaderStyle-Width="10%"></asp:BoundColumn>
										<asp:BoundColumn DataField="interDate1" HeaderText="介紹日期" HeaderStyle-Width="10%"></asp:BoundColumn>
										<asp:BoundColumn DataField="orgName" HeaderText="求才單位" HeaderStyle-Width="15%"></asp:BoundColumn>
										<asp:BoundColumn DataField="workName" HeaderText="職稱名稱" HeaderStyle-Width="15%"></asp:BoundColumn>
										<asp:BoundColumn DataField="workYN" HeaderText="僱用與否" HeaderStyle-Width="10%"></asp:BoundColumn>
										<asp:BoundColumn DataField="NRReason" HeaderText="僱主未能錄用原因" HeaderStyle-Width="20%"></asp:BoundColumn>
										<asp:BoundColumn DataField="NWReason" HeaderText="求職未能推介原因" HeaderStyle-Width="20%"></asp:BoundColumn>
									</Columns>
									<PagerStyle Visible="False" />
								</asp:DataGrid>
							</td>
						</tr>
					</table>
				</asp:Panel>
			</td>
		</tr>
		<tr>
			<td>
			</td>
		</tr>
		<tr>
			<td>
			</td>
		</tr>
	</table>
	</form>
</body>
</html>
