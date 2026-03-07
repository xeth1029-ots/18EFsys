<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SD_05_009.aspx.vb" Inherits="WDAIIP.SD_05_009" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>鍾點費試算</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../style.css" type="text/css" rel="stylesheet">
		<script language="javascript" src="../../js/common.js"></script>
		<script language="javascript" src="../../js/date-picker.js"></script>
		<script language="javascript" src="../../js/openwin/openwin.js"></script>
		<script language="javascript">
			function print(){
				var msg='';
				if (document.form1.OCIDValue1.value=='') msg+='請選擇班級職類\n';
				if (document.form1.years.selectedIndex==0) msg+='請選擇年度\n';
				if (document.form1.months.selectedIndex==0) msg+='請選擇月份\n';
			
				if (msg!=''){
					alert(msg);
					return false;
				}
			}
		//全選
				function check_choice(ischecked){
					for (var i=0; i<form1.elements.length; i++) {
						if (form1.elements[i].type == "checkbox" && form1.elements[i].name.indexOf("CB_teacherList")==0) {
							form1.elements[i].checked = ischecked;
						}
					}
				}
					
					
			 function check(){
				var msg='';
				if (getCheckBoxListValue('CB_teacherList').toString(10)==0) msg+='請選取講師名稱\n';
				if(msg!=''){
					alert(msg);
					return false;
				}
			}
		</script>
	</HEAD>
	<BODY>
		<form id="form1" method="post" runat="server">
			<asp:panel id="Panel1" style="Z-INDEX: 101; POSITION: absolute; TOP: 168px; LEFT: 240px" runat="server"
				CssClass="font" Visible="False" Width="168px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<asp:label id="msg" runat="server" ForeColor="Red"></asp:label></asp:panel>
			<TABLE id="Table1" cellSpacing="1" cellPadding="1" width="600" border="0">
				<TR>
					<TD>
						<TABLE class="font" id="Table2" cellSpacing="1" cellPadding="1" width="100%" border="0">
							<TR>
								<TD>首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">鍾點費試算</font></TD>
							</TR>
						</TABLE>
						<TABLE class="font" id="Table3" cellSpacing="1" cellPadding="1" width="100%" border="0">
							<TR>
								<TD bgColor="#2aafc0">
									<FONT color="#ffffff">&nbsp;&nbsp; 月份</FONT><FONT color="red">*</FONT></TD>
								<TD bgColor="#ecf7ff"><asp:dropdownlist id="years" runat="server" Width="96px"></asp:dropdownlist><FONT face="新細明體">年</FONT>
									<asp:dropdownlist id="months" runat="server" Width="88px"></asp:dropdownlist><FONT face="新細明體">月</FONT></TD>
							</TR>
							<TR>
								<TD width="100" bgColor="#2aafc0"><font color="#ffffff">&nbsp;&nbsp; 職類/班別</font><font color="#ff0000">*</font></TD>
								<TD bgColor="#ecf7ff"><asp:textbox id="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:textbox><asp:textbox id="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:textbox><INPUT onclick="window.open('../02/SD_02_ch.aspx','','width=540,height=520,location=0,status=0,menubar=0,scrollbars=1,resizable=0');"
										type="button" value="..."><INPUT id="TMIDValue1" style="WIDTH: 35px; HEIGHT: 22px" type="hidden" name="Hidden2"
										runat="server"><INPUT id="OCIDValue1" style="WIDTH: 40px; HEIGHT: 22px" type="hidden" name="Hidden1"
										runat="server">
									<BR>
									<span id="HistoryList" style="POSITION: absolute; DISPLAY: none; LEFT: 270px">
										<asp:Table id="HistoryTable" runat="server" Width="310"></asp:Table></span></TD>
							</TR>
							<TR>
								<TD colSpan="2">
									<p align="center"><FONT face="新細明體"><asp:button id="calculate_button" runat="server" Text="試算講師"></asp:button>&nbsp;</FONT>
										<asp:button id="Button1" runat="server" Text="查詢"></asp:button></p>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			<br>
			<br>
			<br>
			<asp:panel id="Panel" runat="server" CssClass="font" Visible="False" Width="600px" Height="93px">
				<TABLE id="search_tbl" class="font" border="1" cellSpacing="0" cellPadding="0" width="600"
					runat="server">
				</TABLE>
				<P><FONT face="新細明體">請選擇有授課的講師名稱&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<asp:CheckBox id="CB1" runat="server" CssClass="font" Text="全選"></asp:CheckBox></FONT></P>
				<P></P>
				<P>
					<asp:CheckBoxList id="CB_teacherList" runat="server" Width="100%" CssClass="font" RepeatColumns="7"
						RepeatDirection="Horizontal"></asp:CheckBoxList></P>
				<P><FONT face="新細明體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					</FONT>
					<asp:Button id="count_Button" runat="server" Text="試算"></asp:Button></P>
			</asp:panel>
			<P></P>
			<P><FONT face="新細明體"></FONT>&nbsp;</P>
		</form>
	</BODY>
</HTML>
