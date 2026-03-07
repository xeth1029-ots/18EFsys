<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_06_001.aspx.vb" Inherits="WDAIIP.TC_06_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級變更審核</title>
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <asp:HiddenField ID="hidLID" runat="server" />
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級變更審核</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="center" Width="60%" runat="server" onfocus="this.blur()"></asp:TextBox>
                                <input id="Button1" type="button" value="..." name="Button1" runat="server">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                                <input id="Button2" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="Button1" runat="server">
                                <input id="trainValue" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" type="hidden" name="jobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">班級名稱 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="ClassCName" runat="server" MaxLength="50" Columns="44"></asp:TextBox></td>
                            <td class="bluecol" style="width: 20%">期別 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="CyclType" runat="server" Columns="10" MaxLength="2"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">申請日期 </td>
                            <td colspan="3" class="whitecol">
                                <span id="span1" runat="server">
                                    <asp:TextBox ID="ApplySDate" runat="server" onfocus="this.blur()" MaxLength="10" Columns="20"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= ApplySDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                    ～<asp:TextBox ID="ApplyEDate" runat="server" onfocus="this.blur()" MaxLength="10" Columns="20"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= ApplyEDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr id="trReviseStatus2854" runat="server">
                            <td class="bluecol">審核狀態 </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="rblReviseStatus" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="X" Selected="True">未審核</asp:ListItem>
                                    <asp:ListItem Value="O">已審核</asp:ListItem>
                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trRBListExpType" runat="server">
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol">
                                <div align="center">
                                    <asp:Button ID="But_Search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="BtnExp1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="PlanReviseList" Width="100%" runat="server" AutoGenerateColumns="False" CssClass="font" AllowPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <%--<asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫" HeaderStyle-Width="9%"></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="PlanYear" HeaderText="年度計畫" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSNAME2" HeaderText="班級名稱" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDATE" HeaderText="開訓日期" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="FDDATE" HeaderText="結訓日期" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="Address" HeaderText="地址" HeaderStyle-Width="9%"></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="ContactName" HeaderText="聯絡人" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Phone" HeaderText="電話" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="變更項目" HeaderStyle-Width="9%">
                                            <ItemTemplate>
                                                <asp:Label ID="labChgTypeN1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="申請日期" HeaderStyle-Width="9%">
                                            <ItemTemplate>
                                                <asp:Label ID="labApplyDate" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="ONLINESENDDATE" HeaderText="線上送件時間" HeaderStyle-Width="9%"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="備註" HeaderStyle-Width="9%">
                                            <ItemTemplate>
                                                <asp:Label ID="labMemoD1" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--<asp:BoundColumn Visible="False" DataField="SubSeqNO"></asp:BoundColumn>--%>
                                        <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="10%">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Button ID="btnView1" runat="server" Text="檢視" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btnUpdat1" runat="server" Text="修改" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btnEdit" runat="server" Text="審核" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btnDelete" runat="server" Text="刪除" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="btnPartReduc" runat="server" Text="還原" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red" Width="100%"></asp:Label>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
