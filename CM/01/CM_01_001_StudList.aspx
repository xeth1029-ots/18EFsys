<%@ Page Language="vb" AutoEventWireup="false" Codebehind="CM_01_001_StudList.aspx.vb" Inherits="WDAIIP.CM_01_001_StudList" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>調整班級學員名單</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../css/style.css" type="text/css" rel="stylesheet">
		<script>
			function CheckData3(){
				var MyTable=document.getElementById('DAtaGrid3');
				var msg='';
				for(i=1;i<MyTable.rows.length;i++){
					if(MyTable.rows(i).cells(4).children(0).checked){
						if(MyTable.rows(i).cells(5).children(0).selectedIndex==0){
							msg+='學員['+MyTable.rows(i).cells(0).innerHTML+']:'+MyTable.rows(i).cells(1).innerHTML+'設未設定主要參訓身分別\n';	
						}
					}
				}
				
				if(msg!=''){
					alert(msg);
					return false;
				}
			}
		</script>
	</HEAD>
	<body>
		<form id="form1" method="post" runat="server">
			<TABLE id="ListMode1" cellSpacing="1" cellPadding="1" width="740" border="0" runat="server">
				<TR>
					<TD>
						<TABLE class="font" id="Table1" cellSpacing="1" cellPadding="1" width="100%" border="0">
							<TR>
								<TD><FONT face="新細明體">學員名單(僅有非離退訓者)</FONT></TD>
							</TR>
							<TR>
								<TD align="center"><asp:datagrid id="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%">
										<AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />    
                                        <Columns>
											<asp:BoundColumn DataField="StudentID" HeaderText="學號"></asp:BoundColumn>
											<asp:BoundColumn DataField="Name" HeaderText="學員"></asp:BoundColumn>
											<asp:TemplateColumn HeaderText="就安公費">
												<ItemTemplate>
													<INPUT id="Radio2" type="radio" value="Radio2" runat="server">
												</ItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn HeaderText="就安自費">
												<ItemTemplate>
													<INPUT id="Radio1" type="radio" value="Radio1" runat="server">
												</ItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn HeaderText="就保公費">
												<ItemTemplate>
													<INPUT id="Radio4" type="radio" value="Radio4" runat="server">
												</ItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn HeaderText="就保自費">
												<ItemTemplate>
													<INPUT id="Radio3" type="radio" value="Radio3" runat="server">
												</ItemTemplate>
											</asp:TemplateColumn>
										</Columns>
									</asp:datagrid><asp:button id="Button1" runat="server" Text="儲存後關閉" 
                                        CssClass="asp_button_M"></asp:button></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			<TABLE id="ListMode2" cellSpacing="1" cellPadding="1" width="740" border="0" runat="server">
				<TR>
					<TD>
						<TABLE class="font" id="Table3" cellSpacing="1" cellPadding="1" width="100%" border="0">
							<TR>
								<TD><FONT face="新細明體">學員名單(僅有非離退訓者)</FONT></TD>
							</TR>
							<TR>
								<TD align="center"><asp:datagrid id="DataGrid2" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%">
										<AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />    										
                                        <Columns>
											<asp:BoundColumn HeaderText="學號"></asp:BoundColumn>
											<asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
											<asp:TemplateColumn HeaderText="預算別">
												<ItemTemplate>
													<asp:RadioButtonList id="BudgetID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
														<asp:ListItem Value="02">就安</asp:ListItem>
														<asp:ListItem Value="03">就保</asp:ListItem>
													</asp:RadioButtonList>
												</ItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn HeaderText="第一單元(12h)">
												<ItemTemplate>
													<INPUT id="RelClass_Unit1" type="checkbox" runat="server">
												</ItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn HeaderText="第二單元(18h)">
												<ItemTemplate>
													<INPUT id="RelClass_Unit2" type="checkbox" runat="server">
												</ItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn HeaderText="第三單元(6h)">
												<ItemTemplate>
													<INPUT id="RelClass_Unit3" type="checkbox" runat="server">
												</ItemTemplate>
											</asp:TemplateColumn>
										</Columns>
									</asp:datagrid><asp:button id="Button3" runat="server" Text="儲存後關閉" 
                                        CssClass="asp_button_M"></asp:button></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			<TABLE id="ListMode3" cellSpacing="1" cellPadding="1" width="740" border="0" runat="server">
				<TR>
					<TD>
						<TABLE class="font" id="ListTable2" cellSpacing="1" cellPadding="1" width="100%" border="0"
							runat="server">
							<TR>
								<TD><FONT face="新細明體">學員名單(僅有非離退訓者)</FONT></TD>
							</TR>
							<TR>
								<TD align="center">
                                    <asp:datagrid id="DataGrid3" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%">										
										<AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />    
                                        <Columns>
											<asp:BoundColumn HeaderText="學號">
												<HeaderStyle Width="30px"></HeaderStyle>
											</asp:BoundColumn>
											<asp:BoundColumn DataField="Name" HeaderText="姓名">
												<HeaderStyle Width="70px"></HeaderStyle>
											</asp:BoundColumn>
											<asp:BoundColumn HeaderText="一般對象">
												<HeaderStyle HorizontalAlign="Center" Width="60px"></HeaderStyle>
												<ItemStyle HorizontalAlign="Center"></ItemStyle>
											</asp:BoundColumn>
											<asp:BoundColumn HeaderText="特定對象">
												<HeaderStyle HorizontalAlign="Center" Width="60px"></HeaderStyle>
												<ItemStyle HorizontalAlign="Center"></ItemStyle>
											</asp:BoundColumn>
											<asp:TemplateColumn HeaderText="修正">
												<HeaderStyle HorizontalAlign="Center" Width="25px"></HeaderStyle>
												<ItemStyle HorizontalAlign="Center"></ItemStyle>
												<ItemTemplate>
													<INPUT id="Correct" type="checkbox" name="Correct" runat="server"> <INPUT id="MIdentityID1" type="hidden" runat="server">
												</ItemTemplate>
											</asp:TemplateColumn>
											<asp:TemplateColumn HeaderText="主要參訓身分別">
												<ItemTemplate>
													<asp:DropDownList id="MIdentityID" runat="server"></asp:DropDownList>
												</ItemTemplate>
											</asp:TemplateColumn>
										</Columns>
									</asp:datagrid><asp:button id="Button2" runat="server" Text="儲存後關閉" 
                                        CssClass="asp_button_M"></asp:button></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
