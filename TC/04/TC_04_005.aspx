<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_04_005.aspx.vb" Inherits="WDAIIP.TC_04_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級未檢送資料註記</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //選擇全部
        function SelectAll(obj, hidobj) {
            //debugger;
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級未檢送資料註記</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" width="100%">
                        <tr id="trOrg" runat="server">
                            <td class="bluecol" width="16%">&nbsp;訓練機構</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" /><br>
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabTMID" runat="server">&nbsp;訓練職類</asp:Label>
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TB_career_id" runat="server" onfocus="this.blur()" Columns="30" Width="66%"></asp:TextBox>
                                <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="btu_sel" runat="server">
                                <input id="TPlanid" type="hidden" name="TPlanid" runat="server">
                                <input id="trainValue" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" type="hidden" name="jobValue" runat="server">
                                <%--<asp:Button ID="Button2" runat="server" Visible="False" Text="Button2" CssClass="asp_button_M"></asp:Button>--%>
                            </td>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">&nbsp;通俗職類</asp:Label>
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="66%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="16%">&nbsp;班別名稱</td>
                            <td class="whitecol" width="34%">
                                <asp:TextBox ID="ClassName" runat="server" Columns="30" Width="50%"></asp:TextBox>
                            </td>
                            <td class="bluecol" width="16%">&nbsp;期別</td>
                            <td class="whitecol" width="34%">
                                <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="2" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;申請日期</td>
                            <td class="whitecol">
                                <span runat="server">
                                    <asp:TextBox ID="UNIT_SDATE" runat="server" Width="35%" MaxLength="10" ToolTip="日期格式:yyyy/MM/dd"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= UNIT_SDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~
        							<asp:TextBox ID="UNIT_EDATE" runat="server" Width="35%" MaxLength="10" ToolTip="日期格式:yyyy/MM/dd"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= UNIT_EDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                            <td class="bluecol">&nbsp;開訓日期</td>
                            <td class="whitecol">
                                <span runat="server">
                                    <asp:TextBox ID="start_date" Width="35%" runat="server" MaxLength="10" ToolTip="日期格式:yyyy/MM/dd"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
		        					<asp:TextBox ID="end_date" Width="35%" runat="server" MaxLength="10" ToolTip="日期格式:yyyy/MM/dd"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;計畫範圍</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="OrgKind2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="G">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <%--<tr id="tr_AppStage_TP28" runat="server"><td class="bluecol">申請階段</td><td class="whitecol" colspan="3"><asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td></tr>--%>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="AppStage2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font"></asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;審核類型</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="PlanMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="S" Selected="True">審核中</asp:ListItem>
                                    <asp:ListItem Value="Y">已通過</asp:ListItem>
                                    <asp:ListItem Value="R">退件修正(含不通過的)</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;檢送資料</td>
                            <td colspan="3" class="whitecol"><%--檢送資料-未檢送--%>
                                <asp:CheckBox ID="CB_DataNotSent_SCH" runat="server" Text="未檢送資料" ToolTip="(勾選)只有未檢送資料" />
                            </td>
                        </tr>
                    </table>
                    <%-- <table class="table_sch" id="TRA" runat="server" width="100%">
                        <tr>
                            <td width="16%" class="bluecol">&nbsp;已通過審核功能</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="AdvanceMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="S" Selected="True">審核狀態</asp:ListItem>
                                    <asp:ListItem Value="C">取消審核</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>--%>
                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td colspan="4" class="whitecol" align="center">
                                <div align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </div>
                                <br />
                                <div align="center">
                                    <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                </div>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="Labmsg1" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:DataGrid ID="dgPlan" runat="server" CssClass="font" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="3%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanYear_ROC" HeaderText="計畫年度" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="4%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="AppliedDate" HeaderText="申請日期" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="訓練起日" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="FDDate" HeaderText="訓練迄日" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="5%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="DISTNAME" HeaderText="管控單位" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="7%"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="機構名稱" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="7%"></asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="OCID" HeaderText="課程代碼" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="4%"></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="CLASSNAME" HeaderText="班名" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="7%"></asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="DataNotSent" HeaderText="未檢送資料" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="4%"></asp:BoundColumn>--%>
                                        <asp:TemplateColumn HeaderText="未檢送資料" HeaderStyle-Width="4%" ItemStyle-HorizontalAlign="Center">
                                            <ItemTemplate>
                                                <asp:CheckBox ID="CB_DataNotSent" runat="server" />
                                                <asp:HiddenField ID="Hid_PCS" runat="server" />
                                                <asp:HiddenField ID="Hid_OCID" runat="server" />
                                                <asp:HiddenField ID="HID_DataNotSent" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 31px" align="center">
                                <uc1:PageControler ID="Pagecontroler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="Labmsg2" runat="server" ForeColor="red"></asp:Label>
                            </td>
                        </tr>

                    </table>

                </td>
            </tr>
        </table>
    </form>
</body>
</html>
