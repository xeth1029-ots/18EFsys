<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SD_15_013.aspx.vb" Inherits="TIMS.SD_15_013" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SD_15_013</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../style.css" type="text/css" rel="stylesheet">
		<script language="javascript" src="../../js/openwin/openwin.js"></script>
		<script language="javascript" src="../../js/date-picker.js"></script>
		<script language="javascript" src="../../js/common.js"></script>
		<script language="javascript">
		/*	
		function SelectAll(obj,hidobj) {
		var num=getCheckBoxListValue(obj).length;
		var myallcheck=document.getElementById(obj+'_'+0);
				
	    if (document.getElementById(hidobj).value!=getCheckBoxListValue(obj).charAt(0)){
					document.getElementById(hidobj).value=getCheckBoxListValue(obj).charAt(0);
					for(var i=1;i<num;i++){
						var mycheck=document.getElementById(obj+'_'+i);
						mycheck.checked=myallcheck.checked;
					}
				}
		}*/
		function choose_class(){
				document.getElementById('OCID1').value='';
				document.getElementById('TMID1').value='';
				document.getElementById('OCIDValue1').value='';
				document.getElementById('TMIDValue1').value='';
				
				openClass('../02/SD_02_ch.aspx?&RID='+document.getElementById('RIDValue').value);
			}
			
		function SelectAll(obj,hidobj){
				var num=getCheckBoxListValue(obj).length;
				var myallcheck=document.getElementById(obj+'_'+0);
				
				if (document.getElementById(hidobj).value!=getCheckBoxListValue(obj).charAt(0)){
					document.getElementById(hidobj).value=getCheckBoxListValue(obj).charAt(0);
					for(var i=1;i<num;i++){
						var mycheck=document.getElementById(obj+'_'+i);
						mycheck.checked=myallcheck.checked;
					}
				}
			}
		
		//alert(num);
		//alert(myallcheck);
		
		// if (document.getElementById(obj+'_'+0).checked == true)
		// {
		// alert(num);
		//  for (var i=1;i<num;i++)
		//  {
		//   document.getElementById(obj+'_'+i).checked = true;
		//  }
		  
		// }
		//if (document.getElementById(obj+'_'+0).checked == false )
		//{
		//  for (var i=1;i<num;i++)
		//  {
		//   document.getElementById(obj+'_'+i).checked = false;
		 // }
		
		//}
	//}
		</script>
	</HEAD>
	<body MS_POSITIONING="FlowLayout">
		<form id="form1" method="post" runat="server">
			<TABLE class="font" cellSpacing="1" cellPadding="1" width="600" border="0">
				<TR>
					<TD>
						<TABLE class="font" id="Table1" cellSpacing="1" cellPadding="1" width="100%" border="0">
							<TR>
								<TD><asp:label id="TitleLab1" runat="server"></asp:label><asp:label id="TitleLab2" runat="server">
									首頁&gt;&gt;學員動態管理&gt;&gt;產學訓統計表&gt;&gt;<FONT color="#800000">綜合查詢統計表</FONT>
									</asp:label></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD>
						<TABLE class="font" cellSpacing="1" cellPadding="1" width="100%" border="1" runat="server">
							<TR>
								<TD class="SD_TD1" width="100">轄區
								</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="Distid" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" RepeatColumns="3"></asp:checkboxlist><INPUT id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server"></TD>
							</TR>
						    <%--
							<TR>
								<TD class="SD_TD1">訓練機構
								</TD>
								<TD class="SD_TD2"><asp:textbox id="center" runat="server" Width="390px"></asp:textbox><INPUT id="RIDValue" type="hidden" size="1" name="RIDValue" runat="server"><INPUT id="Button2" type="button" value="..." name="Button2" runat="server">
									<asp:button id="Button3" style="DISPLAY: none" runat="server" Text="Button3"></asp:button><br>
									<span id="HistoryList2" style="DISPLAY: none; POSITION: absolute" onclick="GETvalue()">
										<asp:table id="HistoryRID" runat="server" Width="310px"></asp:table></span></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">職類/班別
								</TD>
								<TD class="SD_TD2"><asp:textbox id="TMID1" runat="server" onfocus="this.blur()" width="200px"></asp:textbox><asp:textbox id="OCID1" runat="server" onfocus="this.blur()" width="200px"></asp:textbox><INPUT onclick="choose_class()" type="button" value="...">
									<INPUT id="OCIDValue1" type="hidden" size="1" name="OCIDValue1" runat="server"> <INPUT id="TMIDValue1" type="hidden" size="1" name="TMIDValue1" runat="server">
									<BR>
									<span id="HistoryList" style="DISPLAY: none; LEFT: 270px; POSITION: absolute">
										<asp:table id="HistoryTable" runat="server" Width="310"></asp:table></span></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">計畫範圍
								</TD>
								<TD class="SD_TD2"><asp:radiobuttonlist id="Plankind" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" Width="384px">
										<asp:ListItem Value="0">不區分</asp:ListItem>
										<asp:ListItem Value="G">產業人才投資計畫</asp:ListItem>
										<asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
									</asp:radiobuttonlist></TD>
							</TR>
							--%>
							<TR>
								<TD class="SD_TD1">包班種類
								</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="PackageType" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal"
										RepeatColumns="7"></asp:checkboxlist><INPUT id="PackageHidden" type="hidden" value="0" name="PackageHidden" runat="server"></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">辦訓地縣市
								</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="Tcitycode" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" RepeatColumns="7"></asp:checkboxlist><INPUT id="TcityHidden" type="hidden" value="0" name="TcityHidden" runat="server"></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">立案地縣市
								</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="Ocitycode" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" RepeatColumns="7"></asp:checkboxlist><INPUT id="OcityHidden" type="hidden" value="0" name="OcityHidden" runat="server"></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">訓練業別
								</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="GovClassName" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal"
										RepeatColumns="9"></asp:checkboxlist><INPUT id="GovClassHidden" type="hidden" value="0" name="GovClassHidden" runat="server"></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">訓練職能</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="CCID" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" RepeatColumns="3"></asp:checkboxlist><INPUT id="CCIDHidden" type="hidden" value="0" name="CCIDHidden" runat="server"></TD>
							</TR>
							<TR id="KID_6_TR" runat="server">
								<TD class="SD_TD1">新興產業</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="KID_6" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" RepeatColumns="5"></asp:checkboxlist><INPUT id="KID_6_hid" type="hidden" value="0" name="HID_DepID_6" runat="server">
								</TD>
							</TR>
							<TR id="KID_10_TR" runat="server">
								<TD class="SD_TD1">重點服務業</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="KID_10" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" RepeatColumns="4"></asp:checkboxlist><INPUT id="KID_10_hid" type="hidden" value="0" name="HID_DepID_6" runat="server">
								</TD>
							</TR>
							<TR id="KID_4_TR" runat="server">
								<TD class="SD_TD1">新興智慧型產業</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="KID_4" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" RepeatColumns="5"></asp:checkboxlist><INPUT id="KID_4_hid" type="hidden" value="0" name="HID_DepID_6" runat="server">
								</TD>
							</TR>
							<TR id="TR_7dep" style="DISPLAY: none" runat="server">
								<TD class="SD_TD1">七項重點服務業</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="KID_7" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" RepeatColumns="3"></asp:checkboxlist><INPUT id="KID_7_hid" type="hidden" value="0" name="HID_DepID_6" runat="server">
								</TD>
							</TR>
							<TR>
								<TD class="SD_TD1">是否為學分班
								</TD>
								<TD class="SD_TD2"><asp:radiobuttonlist id="PointYN" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal"></asp:radiobuttonlist></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">是否核定
								</TD>
								<TD class="SD_TD2"><asp:radiobuttonlist id="Apppass" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal"></asp:radiobuttonlist></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">是否結訓
								</TD>
								<TD class="SD_TD2"><asp:radiobuttonlist id="Endclass" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal"></asp:radiobuttonlist></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">是否撥款
								</TD>
								<TD class="SD_TD2"><asp:radiobuttonlist id="Appmoney" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" Width="140px">
										<asp:ListItem Value="A">不區分</asp:ListItem>
										<asp:ListItem Value="1">是</asp:ListItem>
										<asp:ListItem Value="0">否</asp:ListItem>
									</asp:radiobuttonlist></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">是否停辦
								</TD>
								<TD class="SD_TD2"><asp:radiobuttonlist id="Stopclass" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal"></asp:radiobuttonlist></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">開訓日期
								</TD>
								<TD class="SD_TD2"><asp:textbox id="SDate1" runat="server" Width="80px"></asp:textbox><IMG style="CURSOR: hand" onclick="javascript:show_calendar('<%= SDate1.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align=top width="24" height="24" >
									&nbsp;~&nbsp;<asp:textbox id="SDate2" runat="server" Width="80px"></asp:textbox>
									<IMG style="CURSOR: hand" onclick="javascript:show_calendar('<%= SDate2.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align=top width="24" height="24" ></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">結訓日期
								</TD>
								<TD class="SD_TD2"><asp:textbox id="EDate1" runat="server" Width="80px"></asp:textbox><IMG style="CURSOR: hand" onclick="Javascript:show_calendar('<%= EDate1.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align=top width="24" height="24" >
									&nbsp;~&nbsp;<asp:textbox id="EDate2" runat="server" Width="80px"></asp:textbox><IMG style="CURSOR: hand" onclick="Javascript:show_calendar('<%= EDate2.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align=top width="24" height="24" ></TD>
							</TR>
							<TR>
								<TD class="SD_TD1">匯出欄位
								</TD>
								<TD class="SD_TD2"><asp:checkboxlist id="ChbExit" runat="server" Font-Size="X-Small" RepeatDirection="Horizontal" RepeatColumns="3"></asp:checkboxlist><INPUT id="ChbExitHidden" type="hidden" value="0" name="ChbExitHidden" runat="server"></TD>
							</TR>
							<TR>
								<TD align="center" width="100%" colSpan="2"><asp:button id="BtnExp" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:button></TD>
							</TR>
							<TR>
								<TD colSpan="3"><FONT color="#ff0000">匯出欄位說明:<br>
										1. 實際開訓人次：排除不開班及離退訓，學員資料確認<BR>
										2. 結訓人次：排除不開班及離退訓，學員資料確認，班級結訓，結訓成績登錄功能有選擇是否有取得學分資格<BR>
										3. 撥款人次：排除不開班及離退訓，學員資料確認，班級結訓，結訓成績登錄功能有選擇是否有取得學分資格<BR>
										&nbsp;&nbsp; ，補助撥款功能為通過<BR>
										4. [人時成本]計算公式：( 核定補助費 / 核定人次 ) / 訓練時數</FONT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
