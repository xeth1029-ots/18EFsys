<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SV_01_003_Insert.aspx.vb" Inherits="WDAIIP.SV_01_003_Insert" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>問卷視窗</title>
	<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
	<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
	<meta content="JavaScript" name="vs_defaultClientScript">
	<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	<link href="../../css/style.css" type="text/css" rel="stylesheet">
	<script language="JavaScript" src="../../js/common.js"></script>
	<script language="JavaScript">
		function RtnFAnswer(obj2, obj3) {
			//debugger;
			if (!isEmpty(document.getElementById(obj3))) {
				var saValue = document.getElementById(obj3).value.split("|");
				//document.getElementById(obj2).value=document.getElementById(obj3).value;
				document.getElementById(obj2).value = saValue[0];
			}
		}
		//檢查問卷答案
		function CheckDescData(obj1, obj2, obj3) {
			var msg = '';
			if (isEmpty(document.getElementById(obj3))) {
				if (document.getElementById(obj1).value == '') msg += '請輸入【問卷答案】\n';
			}

			if (document.getElementById(obj2).value == '') msg += '請輸入【序號】\n';
			else if (!isUnsignedInt(document.getElementById(obj2).value)) msg += '【序號】必須為數字\n';
			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		//檢查問卷答案
		function CheckDescDataE(obj1, obj2) {
			var msg = '';
			if (document.getElementById(obj1).value == '') msg += '請輸入【問卷答案】\n';

			if (document.getElementById(obj2).value == '') msg += '請輸入【序號】\n';
			else if (!isUnsignedInt(document.getElementById(obj2).value)) msg += '【序號】必須為數字\n';

			if (msg != '') {
				alert(msg);
				return false;
			}
		}

		function CheckData() {
			var msg = '';
			if (document.getElementById('Question').value == '') { msg += '請輸入【問卷題目】\n'; }
			if (document.getElementById('SerialQ').value == '') { msg += '請輸入【題目排序序號】\n'; }
			else if (!isUnsignedInt(document.getElementById('SerialQ').value)) { msg += '【題目排序序號】必須為數字\n'; }
			if (document.getElementById('Answercount').value == 'N') { msg += '請輸入【問卷答案】\n'; }
			if (msg != '') {
				alert(msg);
				return false;
			}
		}
		 
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<asp:Panel ID="Panel1" runat="server">
		<asp:PlaceHolder ID="PlaceHolder1" runat="server"></asp:PlaceHolder>
	</asp:Panel>
	<table width="100%">
		<tr width="100%">
			<td align="center">
				<input id="RETop" type="button" value="回上一頁" name="RETop" runat="server" class="button_b_M" >
			</td>
		</tr>
	</table>
	<table class="font" id="Table_I" cellspacing="1" cellpadding="1" width="740" border="0" runat="server">
		<tr>
			<td class="bluecol" width="100">
				<label>
					問卷類別標題</label>
			</td>
			<td class="td_light">
				<label id="QLabel" runat="server">
				</label>
			</td>
		</tr>
		<tr>
			<td class="bluecol">
				<label>
					問卷題目</label>&nbsp;
			</td>
			<td class="td_light">
				<font face="新細明體">
					<asp:TextBox ID="Question" runat="server" Width="528px"></asp:TextBox></font>
			</td>
		</tr>
		<tr>
			<td class="bluecol">
				<label>
					問卷項目類別</label>
			</td>
			<td class="td_light">
				<asp:RadioButtonList ID="QTYPE" runat="server" Width="104px" Font-Size="X-Small" RepeatDirection="Horizontal">
					<asp:ListItem Value="1" Selected="True">單選</asp:ListItem>
					<asp:ListItem Value="2">複選</asp:ListItem>
				</asp:RadioButtonList>
			</td>
		</tr>
		<tr>
			<td class="bluecol">
				<label>
					題目排序序號</label>
			</td>
			<td class="td_light">
				<asp:TextBox ID="SerialQ" runat="server" Width="24px"></asp:TextBox>
			</td>
		</tr>
		<%--</TABLE>
			<Table id="DataGridT" runat="server" Width="100%">--%>
		<tr runat="server">
			<td colspan="2">
				<asp:DataGrid ID="DataGrid2" runat="server" Width="100%" Font-Size="X-Small" AutoGenerateColumns="False" ShowFooter="True">
					<AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
					<HeaderStyle CssClass="head_navy"></HeaderStyle>
					<Columns>
						<asp:TemplateColumn HeaderText="序號">
							<ItemTemplate>
								<asp:Label ID="Label1" runat="server">Label</asp:Label>
							</ItemTemplate>
							<FooterTemplate>
								<asp:TextBox ID="FNO1" runat="server" Width="28px"></asp:TextBox>
							</FooterTemplate>
							<EditItemTemplate>
								<asp:TextBox ID="ENO1" runat="server" Width="28px"></asp:TextBox>
							</EditItemTemplate>
						</asp:TemplateColumn>
						<asp:TemplateColumn HeaderText="問卷答案">
							<ItemTemplate>
								<asp:Label ID="LAnswer" runat="server">Label</asp:Label>
							</ItemTemplate>
							<FooterTemplate>
								<asp:TextBox ID="FAnswer" runat="server" Width="288px"></asp:TextBox>
								<asp:DropDownList ID="ddlFAnswer" runat="server">
								</asp:DropDownList>
							</FooterTemplate>
							<EditItemTemplate>
								<asp:TextBox ID="EAnswer" runat="server" Width="288px"></asp:TextBox>
							</EditItemTemplate>
						</asp:TemplateColumn>
						<asp:TemplateColumn HeaderText="功能">
							<ItemTemplate>
								<asp:Button ID="Edit" runat="server" Text="修改" CommandName="Edit"></asp:Button>
								<asp:Button ID="Del" runat="server" Text="刪除" CommandName="Del"></asp:Button>
							</ItemTemplate>
							<FooterTemplate>
								<asp:Button ID="Save2" runat="server" Text="新增" CommandName="Save2"></asp:Button>
							</FooterTemplate>
							<EditItemTemplate>
								<asp:Button ID="Save" runat="server" Text="儲存" CommandName="update"></asp:Button>
								<asp:Button ID="Cancel" runat="server" Text="取消" CommandName="Cancel"></asp:Button>
							</EditItemTemplate>
						</asp:TemplateColumn>
					</Columns>
				</asp:DataGrid>
			</td>
		</tr>
		<tr>
			<td align="center" colspan="2">
				<asp:Button ID="Save_Q" runat="server" Text="儲存問卷題目"></asp:Button>&nbsp;
				<input id="return1" type="button" value="回上一頁" name="return1" runat="server" class="button_b_M">
			</td>
		</tr>
	</table>
	<input id="Answercount" type="hidden" name="Answercount" runat="server">
	<input id="SKID" type="hidden" name="SKID" runat="server">
	<input id="Type" type="hidden" name="Type" runat="server">
	<input id="SQID" type="hidden" name="SQID" runat="server">
	<input id="Ivalue" type="hidden" name="Ivalue" runat="server">
	<asp:HiddenField ID="HID_SVID" runat="server" />
	<asp:HiddenField ID="hidIptName" runat="server" />
	<asp:HiddenField ID="HidSerial1" runat="server" />
	</form>
</body>
</html>
