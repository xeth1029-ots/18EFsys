<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_014.aspx.vb" Inherits="WDAIIP.SD_15_014" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_15_014</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button5').click();
        }
        function print() {
            var msg = '';

            //if(document.getElementById('OCIDValue1').value=='') msg+='請選擇班級\n';
            if (document.getElementById('yearlist').selectedIndex == 0) msg += '請選擇年度\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function choose_class() {
            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';

            openClass('../02/SD_02_ch.aspx?&RID=' + document.getElementById('RIDValue').value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;訓後動態調查表執行比率</asp:Label>
                </td>
            </tr>
        </table>
        <input id="Years" type="hidden" name="Years" runat="server">
        <%--
			<INPUT id="SOCIDValue" type="hidden" name="SOCIDValue" runat="server">
        --%>
        <%--<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td>
                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                <asp:Label ID="TitleLab2" runat="server">
						首頁&gt;&gt;學員動態管理&gt;&gt;統計表(產學訓表單列印)&gt;&gt;<FONT color="#990000">訓後動態調查表執行比率</FONT>
                </asp:Label>
            </td>
        </tr>
    </table>--%>
        <table class="table_nw" width="100%">
            <tr>
                <td class="bluecol" width="20%">年度
                </td>
                <td class="whitecol" colspan="3">
                    <asp:DropDownList ID="yearlist" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">訓練機構
                </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                    <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                    <asp:Button ID="Button5" Style="display: none" runat="server" Text="Button5"></asp:Button>
                    <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%">
                        </asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">職類/班別
                </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                    <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                        <asp:Table ID="HistoryTable" runat="server" Width="100%">
                        </asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">結訓日期
                </td>
                <td class="whitecol">
                    <span id="span01" runat="server">
                        <asp:TextBox ID="FTDate1" runat="server" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="Javascript:show_calendar('<%= FTDate1.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        &nbsp;~&nbsp;<asp:TextBox ID="FTDate2" runat="server" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="Javascript:show_calendar('<%= FTDate2.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </span>
                </td>
            </tr>
            <tr id="trPlanKind" runat="server">
                <td class="bluecol" width="20%">計畫範圍
                </td>
                <td class="whitecol" colspan="4">
                    <asp:RadioButtonList ID="SearchPlan" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="A">不區分</asp:ListItem>
                        <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                        <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr id="trPackageType" runat="server">
                <td class="bluecol" width="20%">包班種類
                </td>
                <td class="whitecol" colspan="4">
                    <asp:RadioButtonList ID="PackageType" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="A" Selected="True">全部</asp:ListItem>
                        <%--<asp:ListItem Value="1">非包班</asp:ListItem>--%>
                        <asp:ListItem Value="2">企業包班</asp:ListItem>
                        <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td>
                    <p align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
        </table>
        <!--
			<TABLE id="DataGridTable" cellSpacing="1" cellPadding="1" width="100%" border="0" runat="server">
				<tr>
					<td>
						<asp:DataGrid id="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
							<Columns>
								<asp:TemplateColumn HeaderText="調查項目"></asp:TemplateColumn>
								<asp:TemplateColumn HeaderText="調查內容"></asp:TemplateColumn>
								<asp:TemplateColumn HeaderText="一般身分者"></asp:TemplateColumn>
							</Columns>
						</asp:DataGrid>
					</td>
				</tr>
			</TABLE>
			-->
    </form>
</body>
</html>
