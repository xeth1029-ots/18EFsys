<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_01_013.aspx.vb" Inherits="WDAIIP.SD_01_013" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>甄試人數查詢</title>
    <meta content="microsoft visual studio .net 7.1" name="generator" />
    <meta content="visual basic .net 7.1" name="code_language" />
    <meta content="javascript" name="vs_defaultclientscript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetschema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;甄試人數查詢</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" style="width: 20%">訓練機構
                </td>
                <td class="whitecol" colspan="3">
                    <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                    <input id="org" type="button" value="..." name="org" runat="server" class="asp_button_Mini" />
                    <span id="HistoryList2" style="position: absolute; display: none">
                        <asp:Table ID="historyrid" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">
                    <asp:Label ID="labtmid" runat="server">訓練職類</asp:Label>
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                    <input id="btu_sel" type="button" value="..." name="btu_sel" class="asp_button_Mini" runat="server" />
                    <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                    <input id="jobValue" type="hidden" name="jobValue" runat="server" />
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="bluecol">
                    <asp:Label ID="labcjob_unkey" runat="server">通俗職類</asp:Label>
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" class="asp_button_Mini" runat="server" />
                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="bluecol">班級名稱
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="tb_classname" runat="server" Columns="30" Width="40%"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓日期
                </td>
                <td class="whitecol">
                    <span id="span1" runat="server">
                        <asp:TextBox ID="start_date" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
				    <asp:TextBox ID="end_date" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <%--<tr>
			<td class="bluecol">
				報名日期
			</td>
			<td class="whitecol">
				<asp:TextBox ID="redate_start" Width="80" onfocus="this.blur()" runat="server"></asp:TextBox>
				<img style="cursor: pointer" onclick="javascript:show_calendar('<%= redate_start.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
				<asp:TextBox ID="redate_end" Width="80" onfocus="this.blur()" runat="server"></asp:TextBox>
				<img style="cursor: pointer" onclick="javascript:show_calendar('<%= redate_end.clientid %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
			</td>
		</tr>
		<tr>
			<td class="bluecol">
				開班狀態
			</td>
			<td class="whitecol">
				<asp:RadioButtonList ID="NotOpen" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font">
					<asp:ListItem Value="N" Selected="true">開班</asp:ListItem>
					<asp:ListItem Value="Y">不開班</asp:ListItem>
				</asp:RadioButtonList>
			</td>
		</tr>--%>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol" align="center">
                    <asp:Label ID="labpagesize" runat="server" ForeColor="slateblue">顯示列數</asp:Label>
                    <asp:TextBox ID="txtpagesize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                    <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td class="whitecol" align="center">
                    <asp:Label ID="msg" runat="server" ForeColor="red" CssClass="font"></asp:Label>
                </td>
            </tr>
        </table>
        <table id="table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <div id="div1" runat="server">
                        <asp:DataGrid ID="DG_Classinfo" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" AllowPaging="true" AllowSorting="true" Visible="false" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5" />
                            <HeaderStyle CssClass="head_navy" />
                            <Columns>
                                <asp:BoundColumn HeaderText="序號">
                                    <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="orgname" HeaderText="訓練機構">
                                    <HeaderStyle HorizontalAlign="Center" ForeColor="#B0E2FF" Width="20%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="OCLASSID" HeaderText="班別代碼">
                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn HeaderText="開結訓日">
                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="classcname" HeaderText="班別名稱">
                                    <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="trainname" HeaderText="訓練職類">
                                    <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="TNum" HeaderText="訓練人數">
                                    <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundColumn>
                                <asp:BoundColumn DataField="StEnTeNum2" HeaderText="甄試人數">
                                    <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center" />
                                </asp:BoundColumn>
                            </Columns>
                            <PagerStyle Visible="false"></PagerStyle>
                        </asp:DataGrid>
                    </div>

                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
        <%--<tr><td align="center"><asp:Button ID="btnSave1" runat="server" Text="儲存" />
	&nbsp;<asp:Button ID="btnBack1" runat="server" Text="取消" /></td>
	</tr> <asp:Label Style="z-index: 0" ID="msg2_28" runat="server" CssClass="font" ForeColor="blue">
	報名人數1：網路報名人數+現場報名人數-報名者自行取消人數-網路審核失敗人數<br />
	報名人數2：網路報名人數+現場報名人數-報名者自行取消人數	</asp:Label>--%>
    </form>
</body>
</html>
