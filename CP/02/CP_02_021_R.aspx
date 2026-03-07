<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CP_02_021_R.aspx.vb" Inherits="WDAIIP.CP_02_021_R" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>CP_02_021_R</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="../../style.css" type="text/css" rel="stylesheet">
		<script language="javascript" src="../../js/date-picker.js"></script>
		<script language="javascript" src="../../js/openwin/openwin.js"></script>
		<script src="../../js/common.js"></script>
		<script language="javascript">		
			function print(){
				var msg='';
				
				if(document.form1.syear.selectedIndex==0) msg+='請選擇年度\n';
				if(document.form1.smonth.selectedIndex==0) msg+='請選擇月份\n';			
				
				if (msg!=''){
					alert(msg);
					return false;
				}
			}
		</script>
	</HEAD>
	<body>
		<form id="form1" method="post" runat="server">
			<TABLE id="Table1" cellSpacing="1" cellPadding="1" width="600" border="0">
				<TR>
					<TD>
						<TABLE class="font" id="Table2" cellSpacing="1" cellPadding="1" width="100%" border="0">
							<TR>
								<TD>
									<asp:label id="TitleLab1" runat="server"></asp:label>
									<asp:label id="TitleLab2" runat="server">
									首頁&gt;&gt;訓練查核與績效管理&gt;&gt;公務統計報表&gt;&gt;<font color="#990000">公訓組－年度各項練人數統計表</font>
									</asp:label>
								</TD>
							</TR>
						</TABLE>
						<TABLE class="font" id="Table3" cellSpacing="1" cellPadding="1" width="100%" border="0">
							<TR>
								<TD width="100" bgColor="#cc6666"><font color="#ffffff">&nbsp;&nbsp;&nbsp; 統計月份</font><FONT color="#ffff80"><STRONG style="FONT-WEIGHT: 400">*</STRONG></FONT></FONT></TD>
								<TD bgColor="#ffecec"><asp:dropdownlist id="syear" runat="server" AutoPostBack="True"></asp:dropdownlist><FONT face="新細明體">年</FONT>
									<asp:dropdownlist id="smonth" runat="server"></asp:dropdownlist><FONT face="新細明體">月</FONT></TD>
							</TR>
							<TR>
								<TD colSpan="2">
									<p align="center"><asp:button id="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:button></p>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
