<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TrainJob.aspx.vb" Inherits="WDAIIP.TrainJob" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>請選擇職類</title>
	<meta content="microsoft visual studio .net 7.1" name="generator" />
    <meta content="visual basic .net 7.1" name="code_language" />
    <meta content="javascript" name="vs_defaultclientscript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetschema" />
    <link href="../css/style.css" type="text/css" rel="stylesheet">
	<script type="text/javascript" language="javascript">
		function ReturnValue() {
			//debugger;
			var Hid_TPlanIDtype123 = document.getElementById('Hid_TPlanIDtype123');
			var Hid_PERC100 = document.getElementById('Hid_PERC100');
			if (Hid_TPlanIDtype123.value == '1') {
				//tims
				if (document.form1.bus.value == "") {
					alert('請選擇行業別!!!');
					return false;
				} else if (document.form1.job.value == "") {
					alert('請選擇職業分類!!!');
					return false;
				} else if (document.form1.train.value == "") {
					alert('請選擇訓練職類!!!');
					return false;
				}
				opener.document.form1.elements[document.form1.fieldname.value].value = form1.train.options[form1.train.selectedIndex].text;
				opener.document.form1.trainValue.value = document.form1.train.value;
			}
			if (Hid_TPlanIDtype123.value == '2') {
				//tims28
				if (document.form1.bus.value == "") {
					alert('請選擇行業別!!!');
					return false;
				} else if (document.form1.job.value == "") {
					alert('請選擇訓練業別!!!');
					return false;
				}
				opener.document.form1.elements[document.form1.fieldname.value].value = form1.job.options[form1.job.selectedIndex].text;
				opener.document.form1.jobValue.value = document.form1.job.value;
			}
			if (Hid_TPlanIDtype123.value == '3') {
				//tims
				if (document.form1.bus.value == "") {
					alert('請選擇 支用標準!!!');
					return false;
				} else if (document.form1.job.value == "") {
					alert('請選擇 職類課程!!!');
					return false;
				} else if (document.form1.train.value == "") {
					alert('請選擇 業別!!!');
					return false;
				}
				opener.document.form1.elements[document.form1.fieldname.value].value = form1.train.options[form1.train.selectedIndex].text;
				opener.document.form1.trainValue.value = document.form1.train.value; //TMID
				if (opener.document.form1.jobValue) { opener.document.form1.jobValue.value = document.form1.job.value; } //JOBID
				var openerHid_PERC100 = opener.document.form1.Hid_PERC100;
				if (openerHid_PERC100) { openerHid_PERC100.value = Hid_PERC100.value; }
			}
		}
		//top.document.title="請選擇業別";
	</script>
</head>
<body>
	<form id="form1" method="post" runat="server">
	<input id="fieldname" type="hidden" name="fieldname" runat="server">
	<table class="table_nw" border="0" width="100%">
		<tr>
			<td id="busTD" runat="server" class="bluecol" width="20%">行業別 </td>
			<td class="whitecol" width="80%"><asp:DropDownList ID="bus" AutoPostBack="True" runat="server"></asp:DropDownList></td>
		</tr>
		<tr>
			<td id="jobTD" runat="server" class="bluecol" width="20%">職業分類 </td>
			<td class="whitecol" width="80%"><asp:DropDownList ID="job" AutoPostBack="True" runat="server"></asp:DropDownList></td>
		</tr>
		<tr id="trainTR" runat="server">
			<td id="trainTD" runat="server" class="bluecol" width="20%">訓練職類 </td>
			<td class="whitecol" width="80%"><asp:DropDownList ID="train" runat="server"></asp:DropDownList></td>
		</tr>
		<tr>
			<td align="center" colspan="2" class="whitecol"><input type="button" name="but_sub" runat="server" value="選擇" class="asp_button_M" onclick="javascript:ReturnValue();" id="Button1"></td>
		</tr>
	</table>
	<asp:HiddenField ID="Hid_TPlanIDtype123" runat="server" />
	<asp:HiddenField ID="Hid_PERC100" runat="server" />
	</form>
</body>
</html>