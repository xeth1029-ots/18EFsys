<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SD_01_003.aspx.vb" Inherits="TIMS.SD_01_003" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title>sd_01_003</title>
		<meta content="microsoft visual studio .net 7.1" name="generator" />
		<meta content="visual basic .net 7.1" name="code_language" />
		<meta content="javascript" name="vs_defaultclientscript" />
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetschema" />
		<link href="../../css/style.css" type="text/css" rel="stylesheet" />
		<script type="text/javascript" src="../../js/date-picker.js"></script>
		<script type="text/javascript" src="../../js/openwin/openwin.js"></script>
		<script type="text/javascript">
		function choose_class(){				
		    openclass('../02/sd_02_ch.aspx?rid='+document.form1.ridvalue.value);
	    }
		</script>
	</head>
	<body ms_positioning="flowlayout">
		<form id="form1" method="post" runat="server">
			<table class="font" width="740">
				<tr>
					<td class="font">
						<asp:label id="titlelab1" runat="server"></asp:label>
						<asp:label id="titlelab2" runat="server">
						首頁&gt;&gt;學員動態管理&gt;&gt;招生報名&gt;&gt;<font color="#990000">職訓前工作狀況調查表</font>
						</asp:label>
					</td>
				</tr>
			</table>
			<table class="table_nw" width="740">
				<tr>
					<td width="100" class="bluecol">訓練機構</td>
					<td class="whitecol">
						<asp:textbox id="center" runat="server" onfocus="this.blur()" width="310px"></asp:textbox>
                        <input type="button" value="..." id="button2" name="button2" runat="server" 
                            class="button_b_Mini" />
                        <input id="RIDValue" type="hidden" name="hidden2" runat="server" size="1" /><br />
						<span id="HistoryList2" style="display: none;  position: absolute">
						<asp:table id="historyrid" runat="server" width="310px"></asp:table></span></td>
				</tr>
				<tr>
					<td width="100" class="bluecol">職類/班別</td>
					<td class="whitecol">
						<asp:textbox id="TMID1" runat="server" onfocus="this.blur()"></asp:textbox>
                        <asp:textbox id="OCID1" runat="server" onfocus="this.blur()"></asp:textbox>
                        <input onclick="choose_class()" type="button" value="..." 
                            class="button_b_Mini" />
                        <input id="TMIDValue1" style="width: 40px; height: 22px" type="hidden" size="1" name="hidden1" runat="server" />
                        <input id="OCIDValue1" style="width: 32px; height: 22px" type="hidden" size="1" name="hidden3" runat="server" />
						<asp:button id="search" runat="server" text="查詢" CssClass="asp_button_S"></asp:button><br />
						<span id="HistoryList" style="display: none; z-index: 101; left: 270px; position: absolute">
						<asp:table id="historytable" runat="server" width="310px"></asp:table></span></td>
				</tr>
			</table>
			<input id="check_search" style="width: 64px; height: 22px" type="hidden" size="5" name="check_search" runat="server" /> 
            <input id="check_add" style="width: 64px; height: 22px" type="hidden" size="5" name="check_add" runat="server" />
            <input id="check_mod" style="width: 64px; height: 22px" type="hidden" size="5" name="check_mod" runat="server" /> 
            <input id="check_del" style="width: 64px; height: 22px" type="hidden" size="5" name="check_del" runat="server" />
			<br />
			<div>
					<asp:panel id="panel1" runat="server" height="168px" width="600px">
						<p align="center"><asp:label id="msg" runat="server" cssclass="font" forecolor="red"></asp:label></p>						
						<asp:datagrid id="datagrid1" runat="server" width="592px" cssclass="font" autogeneratecolumns="false">
                            <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                            <HeaderStyle CssClass="head_navy" />
							<columns>
								<asp:boundcolumn headertext="序號">
									<headerstyle HorizontalAlign="Center"></headerstyle>
									<itemstyle HorizontalAlign="Center"></itemstyle>
								</asp:boundcolumn>
								<asp:boundcolumn headertext="班別">
									<headerstyle HorizontalAlign="Center"></headerstyle>
									<itemstyle HorizontalAlign="Center"></itemstyle>
								</asp:boundcolumn>
								<asp:boundcolumn datafield="total" headertext="在訓人數">
									<headerstyle HorizontalAlign="Center"></headerstyle>
									<itemstyle HorizontalAlign="Center"></itemstyle>
								</asp:boundcolumn>
								<asp:boundcolumn datafield="num" headertext="填寫人數">
									<headerstyle HorizontalAlign="Center"></headerstyle>
									<itemstyle HorizontalAlign="Center"></itemstyle>
								</asp:boundcolumn>
								<asp:templatecolumn headertext="功能">
									<headerstyle HorizontalAlign="Center"></headerstyle>
									<itemstyle HorizontalAlign="Center"></itemstyle>
									<itemtemplate>
										<asp:LinkButton id="button1" runat="server" text="查詢" commandname="view" CssClass="linkbutton"></asp:LinkButton>
									</itemtemplate>
								</asp:templatecolumn>
								<asp:boundcolumn visible="false" datafield="ftdate" headertext="結訓日期"></asp:boundcolumn>
								<asp:boundcolumn visible="false" datafield="cycltype" headertext="cycltype"></asp:boundcolumn>
								<asp:boundcolumn visible="false" datafield="leveltype" headertext="leveltype"></asp:boundcolumn>
							</columns>
						</asp:datagrid>
						<p>
							<asp:label id="label1" runat="server" cssclass="font"></asp:label></p>
						<p style="margin-top: 3px; margin-bottom: 3px" align="center">
							<asp:label id="msg2" runat="server" cssclass="font" forecolor="red"></asp:label></p>
						<table id="studenttable" cellspacing="1" cellpadding="1" width="500" border="0" runat="server">
							<tr>
								<td>
									<asp:datagrid id="dg_stud" runat="server" width="100%" cssclass="font" autogeneratecolumns="false">
                                        <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                        <HeaderStyle CssClass="head_navy" />
										<columns>
											<asp:boundcolumn headertext="學號">
												<headerstyle HorizontalAlign="Center"></headerstyle>
												<itemstyle HorizontalAlign="Center"></itemstyle>
											</asp:boundcolumn>
											<asp:boundcolumn datafield="name" headertext="姓名(離退訓日期)">
												<headerstyle HorizontalAlign="Center"></headerstyle>
												<itemstyle HorizontalAlign="Center"></itemstyle>
											</asp:boundcolumn>
											<asp:boundcolumn headertext="填寫狀態">
												<headerstyle HorizontalAlign="Center"></headerstyle>
												<itemstyle HorizontalAlign="Center"></itemstyle>
											</asp:boundcolumn>
											<asp:templatecolumn headertext="功能">
												<headerstyle HorizontalAlign="Center"></headerstyle>
												<itemstyle HorizontalAlign="Center"></itemstyle>
												<itemtemplate>
													<asp:LinkButton id="button4" runat="server" text="新增" commandname="insert" CssClass="linkbutton"></asp:LinkButton>
													<asp:LinkButton id="button5" runat="server" text="查詢" commandname="check" CssClass="linkbutton"></asp:LinkButton>
													<asp:LinkButton id="button6" runat="server" text="清除重填" commandname="clear" CssClass="linkbutton"></asp:LinkButton>
												</itemtemplate>
											</asp:templatecolumn>
											<asp:boundcolumn visible="false" datafield="ocid" headertext="ocid"></asp:boundcolumn>
											<asp:boundcolumn visible="false" datafield="socid" headertext="socid"></asp:boundcolumn>
											<asp:boundcolumn visible="false" datafield="studentid" headertext="studentid"></asp:boundcolumn>
										</columns>
									</asp:datagrid></td>
							</tr>
							<tr>
								<td align="center"><font face="新細明體"></font></td>
							</tr>
						</table>
						<p><font face="新細明體"></font>&nbsp;</p>
					</asp:panel></font></div>
		</form>
	</body>
</html>
