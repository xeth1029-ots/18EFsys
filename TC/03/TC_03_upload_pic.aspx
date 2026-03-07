<%@ Page Language="vb" AutoEventWireup="false" Codebehind="TC_03_upload_pic.aspx.vb" Inherits="TIMS.TC_03_upload_pic" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>TC_03_upload_pic</TITLE>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../style.css" type="text/css" rel="stylesheet">
		<script language="javascript" src="../../js/date-picker.js"></script>
		<script language="javascript" src="../../js/common.js"></script>
		<script language="javascript" src="../../js/openwin/openwin.js"></script>
		<script language="JavaScript">
		
			function check_style() {
				if (form1.STDate.disabled) {
					form1.date1.style.cursor = "";
					form1.date1.onclick = null;
				}
				if (form1.FDDate.disabled) {
					form1.date2.style.cursor = "";
					form1.date2.onclick = null;
				}
					
			}
			
			function MM_OpenWindowCenter(owcURL,owcWinName,owcWinWidth,owcWinHeight,owcFeatures) { 
				// owcWinWidth  : screen X  
				// owcWinHeight  : screen Y 
				// owcFeatures:features
				var x=(screen.width-owcWinWidth)/2;
				var y=(screen.height-owcWinHeight)/2;
				window.open(owcURL,owcWinName,"left="+x+", top="+y+", width="+owcWinWidth+", height="+owcWinHeight+owcFeatures);
		    }
		</script>
	</HEAD>
	<body onload="check_style();">
		<FONT face="新細明體"></FONT>
		<form id="form1" method="post" runat="server">
			<table cellSpacing="0" cellPadding="0" width="272" border="0" style="WIDTH: 272px; HEIGHT: 298px">
				<TR>
					<TD class="font" width="187" bgColor="#ffcccc" style="WIDTH: 187px"><FONT>&nbsp;&nbsp;&nbsp;
						</FONT><FONT face="新細明體" class="font">場地圖片上傳</FONT></TD>
					<TD style="WIDTH: 468px">
						<asp:radiobuttonlist id="Radiobuttonlist1" runat="server" CssClass="font" Width="317px" CellSpacing="0"
							CellPadding="0" RepeatDirection="Horizontal">
							<asp:ListItem Value="101" Selected="True">學科教室1</asp:ListItem>
							<asp:ListItem Value="102">學科教室2</asp:ListItem>
							<asp:ListItem Value="201">術科教室1</asp:ListItem>
							<asp:ListItem Value="202">術科教室2</asp:ListItem>
						</asp:radiobuttonlist><INPUT id="File1" type="file" name="File1" runat="server">
						<asp:button id="Button10" runat="server" Text="上傳圖片"></asp:button></TD>
				</TR>
				<tr>
					<TD colspan="2" style="WIDTH: 569px"><FONT face="新細明體">
							<asp:DataGrid id="DataGrid1" runat="server" CssClass="font" AutoGenerateColumns="False" Width="448px"
								AllowPaging="True">
								<ItemStyle BackColor="#ECF7FF"></ItemStyle>
								<HeaderStyle ForeColor="black" BackColor="#ffcccc"></HeaderStyle>
								<Columns>
									<asp:BoundColumn DataField="Index" HeaderText="上傳檔案"></asp:BoundColumn>
									<asp:BoundColumn DataField="FileName" HeaderText="檔案名稱"></asp:BoundColumn>
								</Columns>
								<PagerStyle Visible="False"></PagerStyle>
							</asp:DataGrid></FONT></TD>
				</tr>
				<TR>
					<TD colspan="2" style="WIDTH: 569px; HEIGHT: 23px">
						<P align="center">
							<asp:Button id="Button1" runat="server" Text="關閉此頁"></asp:Button><FONT face="新細明體"></FONT></P>
					</TD>
				</TR>
			</table>
		</form>
		</SCRIPT>
	</body>
</HTML>
