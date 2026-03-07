<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_005_Classid.aspx.vb" Inherits="WDAIIP.TC_01_005_Classid" %>

 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>請選擇隸屬班級代碼</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="javascript" src="../../js/date-picker.js"></script>
	<script language="javascript" src="../../js/openwin/openwin.js"></script>
	<script language="javascript">
		function set_class(clsid, classid, classname) {
			var CLSID = document.getElementById("CLSID");
			var Classid = document.getElementById("Classid");
			var ClassName = document.getElementById("ClassName");
			CLSID.value = clsid;
			Classid.value = classid;
			ClassName.value = classname;
		}

		function returnValue() {
			var msg = document.getElementById("msg");
			var CLSID = document.getElementById("CLSID");
			var Classid = document.getElementById("Classid");
			var ClassName = document.getElementById("ClassName");

			if (msg && msg.innerHTML == "") {
				if (CLSID.value == "") {
					alert("請選擇班別代碼");
					return false;
				}
				else {
					window.close();
				}
			}

			//錯誤訊息為空,表示查有資料	
			//HidOCID1
			//alert(clsid);
			opener.document.form1.Classid.value = Classid.value + '(' + ClassName.value + ')';
			opener.document.form1.Classid_Hid.value = CLSID.value;
			window.close();
		}
		
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<input id="ClassName" type="hidden">
	<input id="CLSID" type="hidden">
	<input id="Classid" type="hidden">
	<table class="table_nw" cellspacing="1" cellpadding="1" width="100%" border="0" >
		<tr>
			<td class="bluecol" style="width:20%">
				年度：
			</td>
			<td id="Td2" align="left"  runat="server" class="whitecol">
				<asp:DropDownList ID="ddlYears" runat="server">
				</asp:DropDownList>
				<font color="red">(2009年是不含年度的舊資料)</font>
			</td>
		</tr>
		<tr>
			<td class="bluecol">
				訓練計畫：
			</td>
			<td id="td1" align="left" runat="server" class="whitecol">
				<asp:DropDownList ID="TPlan_List" runat="server">
				</asp:DropDownList>
			</td>
		</tr>
		<tr>
			<td colspan="5">
				<div align="center" class="whitecol">
					<asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button></div>
			</td>
		</tr>
	</table>
    <div style="overflow-y: auto; height:400px;">
	    <asp:Panel ID="Panel" runat="server" Width="100%">
		    <div align="center">
			    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></div>
		    <asp:DataGrid ID="DG_Class" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" Visible="False" CellPadding="8">
			    <HeaderStyle BackColor="#FFCCCC" CssClass="head_navy"></HeaderStyle>
			    <Columns>
				    <asp:TemplateColumn HeaderText="選取">
					    <HeaderStyle Width="5%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"/>
					    <ItemTemplate>
						    <input type="radio" name="myradio" value="1" id="myradio" runat="server">
					    </ItemTemplate>
				    </asp:TemplateColumn>
				    <asp:BoundColumn DataField="Years" HeaderText="設定年度">
					    <HeaderStyle Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"/>
				    </asp:BoundColumn>
				    <asp:BoundColumn DataField="ClassID" HeaderText="班別代碼">
					    <HeaderStyle Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"/>
				    </asp:BoundColumn>
				    <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱">
                        <HeaderStyle Width="75%"/>
				    </asp:BoundColumn>
			    </Columns>
			    <PagerStyle Visible="False"></PagerStyle>
		    </asp:DataGrid>
		    <p align="center">
			    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
		    </p>
		    <div align="center" class="whitecol">
			    <input id="submit" onclick="returnValue()" type="button" value="確定" runat="server"  Class="asp_button_M"></div>
	    </asp:Panel>
    </div>
	</form>
</body>
</html>
