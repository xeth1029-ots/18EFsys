<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_004_dbl.aspx.vb" Inherits="TIMS.SD_01_004_dbl" %>
<%@ Register TagPrefix="uc1" TagName="PageControler" Src="../../PageControler.ascx" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script language="javascript" type="text/javascript">
        /*
		function printDoc() {
			if (!factory.object) {
				return
			} else {
				factory.printing.header = '';
				factory.printing.footer = '';
				factory.printing.portrait = true;
				factory.printing.Print(true);
			}
		}

		function ShowPersonData(obj) {
			document.getElementById(obj).style.display = 'inline';
		}

		function HidPersonData(obj) {
			document.getElementById(obj).style.display = 'none';
		}
		*/
    </script>
</head>
<body>
    <!-- MeadCo ScriptX -->
    <%--<object style="display: none" id="factory" codebase="../../scriptx/smsx.cab#Version=6,6,440,26" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext>
	</object>--%>
    <form id="form1" runat="server">
        <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="right">
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td align="right">共計：<asp:Label ID="RecordCount" runat="server" />&nbsp;筆資料</td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <table class="Table_nw" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="15%">學員姓名</td>
                            <td class="whitecol" colspan="3"><asp:Label ID="LabName" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">補助費用</td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labMoneyShow1" runat="server"></asp:Label><br />
                                <asp:Label ID="labOver6w" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">補助費用說明</td>
                            <td class="whitecol" colspan="3"><asp:Label ID="labMsg2" runat="server">*預估補助費用是以該課程費用80%作為估算</asp:Label></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center"><asp:Label ID="msgbb" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid2bb" runat="server" Width="100%" AllowSorting="True" AllowPaging="True" PageSize="20" AutoGenerateColumns="False" CssClass="font">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:TemplateColumn>
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <HeaderTemplate>序號</HeaderTemplate>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="Labsignno" runat="server"></asp:Label>
                                    <asp:Label ID="Labdouble" runat="server" ForeColor="Red">重</asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="DistName" HeaderText="轄區">
                                <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="Years" HeaderText="年度">
                                <HeaderStyle Width="5%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位">
                                <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="CLASSCNAME" HeaderText="課程名稱">
                                <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TRound" HeaderText="訓練期間">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn>
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <HeaderTemplate>(重疊)日期<br/>-上課時間</HeaderTemplate>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Literal ID="Literal1" runat="server"></asp:Literal>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn>
                                <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                <HeaderTemplate>訓練狀態</HeaderTemplate>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="Labstudstatus" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center"><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler></td>
            </tr>
            <tr>
                <td align="center"><input id="Btnclose" type="button" value="關閉" runat="server" class="button_b_M"></td>
            </tr>
        </table>
        <%--
        <asp:HiddenField ID="HidDouble" runat="server" />
	    <asp:HiddenField ID="HidMoney6" runat="server" />
        --%>
        <asp:HiddenField ID="Hid_eSerNum" runat="server" />
    </form>
</body>
</html>