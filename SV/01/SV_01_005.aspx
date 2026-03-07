<%@ Register TagPrefix="uc1" TagName="PageControler" Src="../../PageControler.ascx" %>
<%@ Page aspcompat="true" Language="vb" AutoEventWireup="false" Codebehind="SV_01_005.aspx.vb" Inherits="TIMS.SV_01_005" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>職業訓練業務資訊管理網_訓練期末學員滿意度狀況查詢</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../style.css" type="text/css" rel="stylesheet">
		<script language="javascript" src="../../js/date-picker.js"></script>
		<script language="javascript" src="../../js/openwin/openwin.js"></script>
		<script src="../../js/common.js"></script>
		<script language="javascript">
				
		function CheckData(){
				var msg='';
				if(document.form1.IDNO.value=='') msg+='身分證號碼不能是空白\n';
				if(document.form1.birth_date.value=='') msg+='出生日期不能是空白\n';
				if(!IsDate(document.form1.birth_date.value)) msg+='出生日期不是正確的日期格式\n';
				if (msg!=''){
					alert(msg);
					return false;
				}
			}
			function IsDate(MyDate){
				if(MyDate!=''){
					if(!checkDate(MyDate))
						return false;
				}
				return true;
			}
		</script>
	</HEAD>
	<body>
		<form id="form1" method="post" runat="server">
			<table class="font" width="600">
				<tr>
					<td class="font">職業訓練業務資訊管理網&gt;&gt; <FONT color="#990000">訓練期末學員滿意度狀況查詢</FONT>
					</td>
				</tr>
			</table>
			<FONT face="新細明體"></FONT>
			<table class="font" cellSpacing="1" cellPadding="1" width="600" border="0">
				<tr>
					<td id="td6" align="left" width="15%" bgColor="#2aafc0" runat="server">&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#ffffff">身分證號碼<FONT color="red">*</FONT></FONT></td>
					<td width="259" bgColor="#ebf8ff"><asp:textbox id="IDNO" runat="server"></asp:textbox><BR>
						<span id="HistoryList2" style="DISPLAY: none; POSITION: absolute"></span>
					</td>
				</tr>
				<TR>
					<TD id="td5" align="left" bgColor="#2aafc0" runat="server">&nbsp;&nbsp;&nbsp;&nbsp;<FONT color="#ffffff">出生日期<FONT color="red">*</FONT></FONT></TD>
					<TD bgColor="#ebf8ff" colSpan="3"><asp:textbox id="birth_date" Width="80" Runat="server"></asp:textbox><IMG style="CURSOR: hand" onclick="javascript:show_calendar('<%= birth_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align=top width="24" height="24" ></TD>
				</TR>
				<tr>
					<td colSpan="4">
						<DIV align="center"><asp:label id="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:label><asp:textbox id="TxtPageSize" runat="server" Width="23px" MaxLength="2">10</asp:textbox><asp:button id="bt_search" Runat="server" Text="查詢"></asp:button><INPUT id="NY" style="WIDTH: 40px; HEIGHT: 22px" type="hidden" size="1" name="NY" runat="server"><FONT face="新細明體">&nbsp;
							</FONT>
						</DIV>
						<DIV align="center"><FONT face="新細明體"></FONT>&nbsp;</DIV>
						<DIV align="left"><asp:label id="msg" runat="server" ForeColor="Red" CssClass="font"></asp:label></DIV>
					</td>
				</tr>
			</table>
			<TABLE id="Table1" cellSpacing="1" cellPadding="1" width="600" border="0">
			</TABLE>
			<asp:panel id="Panel" runat="server" Width="100%" Visible="False">
				<TABLE class="font" id="search_tbl" cellSpacing="0" cellPadding="0" width="600" border="1"
					runat="server">
				</TABLE>
				<asp:DataGrid id="DG_ClassInfo" runat="server" Width="100%" CssClass="font" Visible="False" AllowSorting="True"
					AllowPaging="True" AutoGenerateColumns="False">
					<AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
					<ItemStyle BackColor="#EBF8FF"></ItemStyle>
					<HeaderStyle ForeColor="White" BackColor="#2AAFC0"></HeaderStyle>
					<Columns>
						<asp:BoundColumn DataField="StudentID" HeaderText="學號">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="Name" HeaderText="姓名">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn HeaderText="管控&lt;br&gt;單位">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="訓練機構">
							<HeaderStyle HorizontalAlign="Center" ForeColor="Black"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="QaySDate" HeaderText="問卷期間起日">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn DataField="QayFDate" HeaderText="問卷期間迄日">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:BoundColumn HeaderText="填寫狀況">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
						</asp:BoundColumn>
						<asp:TemplateColumn HeaderText="功能">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<ItemTemplate>
								<asp:Button id="Edit" runat="server" tooltip="在問卷調查起迄期間內且起迄日期均有值才可新增或修改"></asp:Button>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
						<asp:BoundColumn Visible="False" DataField="StudentID" HeaderText="StudentID"></asp:BoundColumn>
						<asp:BoundColumn Visible="False" DataField="socid" HeaderText="Socid"></asp:BoundColumn>
						<asp:BoundColumn Visible="False" DataField="QuesID" HeaderText="QuesID"></asp:BoundColumn>
					</Columns>
					<PagerStyle Visible="False"></PagerStyle>
				</asp:DataGrid>
				<DIV align="center">
					<UC1:PAGECONTROLER id="PageControler1" runat="server"></UC1:PAGECONTROLER></DIV>
			</asp:panel></form>
	</body>
</HTML>
