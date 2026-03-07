<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="PageControler.ascx.vb" Inherits="WDAIIP.PageControler" TargetSchema="http://schemas.microsoft.com/intellisense/ie5" %>
<script>
	function ChangePage(obj, PageIndex, btn) {
		document.getElementById(obj).value = PageIndex;
		document.getElementById(btn).click();
	}
	function AddPage(Index, obj, btn) {
		var PageIndex = document.getElementById(obj).value;
		PageIndex = parseInt(PageIndex) + Index;
		ChangePage(obj, PageIndex, btn)
	}
	function LastPage(num, obj, btn) {
		if (num == 1)
			ChangePage(obj, 1, btn);
		else
			ChangePage(obj, num, btn);
	}
</script>
<%-- 
<style>
	.LinkSpan
	{
		border-right: #afafaf 1px solid;
		border-top: #afafaf 1px solid;
		font-size: 10pt;
		border-left: #afafaf 1px solid;
		cursor: pointer;
		border-bottom: #afafaf 1px solid;
		text-align: center;
		width: 50px;
		height: 25px;
	}
	
	.LinkSpanImg
	{
		border-right: #afafaf 1px solid;
		border-top: #afafaf 1px solid;
		font-size: 10pt;
		border-left: #afafaf 1px solid;
		cursor: pointer;
		border-bottom: #afafaf 1px solid;
		text-align: center;
		vertical-align: middle;
		width: 20px;
		height: 25px;
	}
	.PageLinkSpan
	{
		border-right: #afafaf 1px solid;
		border-top: #afafaf 1px solid;
		font-size: 10pt;
		border-left: #afafaf 1px solid;
		width: 16px;
		cursor: pointer;
		border-bottom: #afafaf 1px solid;
		height: 1em;
		text-align: center;
	}
	.OverLinkSpan
	{
		border-right: #afafaf 1px solid;
		border-top: #afafaf 1px solid;
		font-size: 10pt;
		border-left: #afafaf 1px solid;
		width: 16px;
		cursor: pointer;
		border-bottom: #afafaf 1px solid;
		height: 1em;
		background-color: #ffff99;
		text-align: center;
	}
</style>
--%>
<table id="PageControlerTable1" border="0" class="font">
	<tr>
		<td>
			<span class="LinkSpan">
				<asp:LinkButton ID="FirstButton" runat="server" CausesValidation="False" ToolTip="最前頁">最前頁</asp:LinkButton></span>
		</td>
		<td>
			<span class="LinkSpan">
				<asp:LinkButton ID="PrePreButton" runat="server" CausesValidation="False" ToolTip="往前跳10頁">往前跳10頁</asp:LinkButton></span>
		</td>
		<td>
			<span class="LinkSpan">
				<asp:LinkButton ID="PreButton" runat="server" ForeColor="Black" CausesValidation="False" ToolTip="上一頁">上一頁</asp:LinkButton></span>
		</td>
		<td>
			<asp:Label ID="PageIndexFlag" runat="server"></asp:Label>
		</td>
		<td>
			<span class="LinkSpan">
				<asp:LinkButton ID="NextButton" runat="server" ForeColor="Black" CausesValidation="False" ToolTip="下一頁">下一頁</asp:LinkButton></span>
		</td>
		<td>
			<span class="LinkSpan">
				<asp:LinkButton ID="NextNextButton" runat="server" CausesValidation="False" ToolTip="往後跳10頁">往後跳10頁</asp:LinkButton></span>
		</td>
		<td>
			<span class="LinkSpan">
				<asp:LinkButton ID="LastButton" runat="server" ImageUrl="images/right2.gif" CausesValidation="False" ToolTip="最後頁">最後頁</asp:LinkButton></span>
		</td>
		<td style="word-wrap: unset;white-space: nowrap;">
			&nbsp;|&nbsp;共：
			<asp:Label ID="PageCountLabel" runat="server"></asp:Label>頁
		</td>
	</tr>
</table>
<input id="NowPage" type="hidden" runat="server">
<input id="HidSSSDTRID" type="hidden" runat="server">
<asp:Button ID="Button1" runat="server" Text="換頁" CausesValidation="False" CssClass="asp_button_M" style="display:none"></asp:Button>
