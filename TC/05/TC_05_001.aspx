<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_05_001.aspx.vb" Inherits="WDAIIP.TC_05_001" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級變更申請</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
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
        if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function SearchMode_CHGACT1() {
            $("#PlanList").hide();
            $("#PageControlerTable1").hide();
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級變更申請</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="Org" type="button" value="..." name="Org" runat="server">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="tr_TMIDVALUE_TP06" runat="server">
                            <td class="bluecol" width="20%">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="TB_career_id" onfocus="this.blur()" runat="server" Columns="30" Width="40%"></asp:TextBox>
                                <input id="Button1" onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." name="Button1" runat="server">
                                <input id="trainValue" type="hidden" name="trainValue" runat="server">
                                <input id="jobValue" type="hidden" name="jobValue" runat="server">
                            </td>
                        </tr>
                        <tr id="tr_CJOBVALUE_TP06" runat="server">
                            <td class="bluecol">
                                <div>
                                    <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                                </div>
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="40%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">班別名稱</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="ClassName" runat="server" MaxLength="50" Columns="44"></asp:TextBox></td>
                            <td class="bluecol" style="width: 20%">期別</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="CyclType" runat="server" MaxLength="2" Columns="10"></asp:TextBox></td>
                        </tr>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">查詢模式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="SearchMode" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="Plan_PlanInfo" Selected="True">申請</asp:ListItem>
                                    <asp:ListItem Value="Plan_Revise">變更結果</asp:ListItem>
                                </asp:RadioButtonList>
                                <%--<asp:RequiredFieldValidator ID="MustSearchMode" runat="server" ControlToValidate="SearchMode" Display="None" ErrorMessage="請選擇查詢模式"></asp:RequiredFieldValidator>--%>
                            </td>
                            <td class="bluecol">審核狀態</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="CheckMode" runat="server">
                                    <asp:ListItem Value="==請選擇==">==請選擇==</asp:ListItem>
                                    <asp:ListItem Value="1">審核不通過</asp:ListItem>
                                    <asp:ListItem Value="2">審核中</asp:ListItem>
                                    <asp:ListItem Value="0">審核完成</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="But_Search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridTable28" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <div>紅色表示此班級已經結訓，不可以提出申請變更([提示]如果想申請變更，請先將班級結訓狀態解除)</div>
                                <div id="divTip" runat="server" visible="false" style="text-align: right;"><font color="#0066FF">※只有「審核中」的資料才提供列印</font></div>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="PlanList28" Width="100%" runat="server" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="PlanYear" HeaderText="年度計畫">
                                            <HeaderStyle Width="8%" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn Visible="False" DataField="PlanName" HeaderText="訓練計畫"><HeaderStyle Width="15%" /></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="ClassName2" HeaderText="班別名稱">
                                            <HeaderStyle Width="18%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類">
                                            <HeaderStyle Width="12%" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="變更項目">
                                            <HeaderStyle Width="12%" />
                                            <ItemTemplate>
                                                <asp:Label ID="labAltDataID" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.AltDataID") %>'></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--asp:TemplateColumn HeaderText="申請變更日"><HeaderStyle Width="10%" /><ItemTemplate><asp:Label ID="labCDate" runat="server" Text='<%# Formatdatetime(DataBinder.Eval(Container, "DataItem.CDate"),2) %>'></asp:Label></ItemTemplate></asp:TemplateColumn--%>
                                        <asp:BoundColumn DataField="CDate" HeaderText="申請變更日">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="REVISEACCT_Name" HeaderText="申請人姓名">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="modifydate" HeaderText="審核時間">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="計畫狀態">
                                            <HeaderStyle Width="10%" />
                                            <ItemTemplate>
                                                <asp:Label ID="labPrjstatus" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="10%" HeaderStyle-HorizontalAlign="Center">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:LinkButton ID="But_Dir" runat="server" Text="申請變更" CausesValidation="false" CssClass="linkbutton" CommandName="appChg"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="線上送件" HeaderStyle-Width="10%" HeaderStyle-HorizontalAlign="Center">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:LinkButton ID="BTN_OL_EDIT1" runat="server" Text="編輯" CausesValidation="false" CssClass="linkbutton" CommandName="OL_EDIT1"></asp:LinkButton>
                                                <asp:LinkButton ID="BTN_OL_SEND1" runat="server" Text="送出" CausesValidation="false" CssClass="linkbutton" CommandName="OL_SEND1"></asp:LinkButton>
                                                <asp:LinkButton ID="BTN_OL_DEL1" runat="server" Text="刪除" CausesValidation="false" CssClass="linkbutton" CommandName="OL_DEL1"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="列印" HeaderStyle-Width="10%" HeaderStyle-HorizontalAlign="Center">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <input id="Button3" type="button" value="計畫變更表" runat="server" class="asp_Export_M" />
                                                <input id="bt_print" type="button" value="變更後課程表" runat="server" class="asp_Export_M" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>紅色表示此班級已經結訓，不可以提出申請變更([提示]如果想申請變更，請先將班級結訓狀態解除)</td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="PlanList" Width="100%" runat="server" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="PlanYear" HeaderText="年度計畫">
                                            <HeaderStyle Width="10%" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn Visible="False" DataField="PlanName" HeaderText="訓練計畫"><HeaderStyle Width="15%" /></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="ClassName2" HeaderText="班別名稱">
                                            <HeaderStyle Width="20%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TrainName" HeaderText="訓練職類">
                                            <HeaderStyle Width="15%" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="變更項目">
                                            <HeaderStyle Width="15%" />
                                            <ItemTemplate>
                                                <asp:Label ID="labAltDataID" runat="server" Text='<%# DataBinder.Eval(Container, "DataItem.AltDataID") %>'></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--asp:TemplateColumn HeaderText="申請變更日"><HeaderStyle Width="10%" /><ItemTemplate><asp:Label ID="labCDate" runat="server" Text='<%# Formatdatetime(DataBinder.Eval(Container, "DataItem.CDate"),2) %>'></asp:Label></ItemTemplate></asp:TemplateColumn--%>
                                        <asp:BoundColumn DataField="CDate" HeaderText="申請變更日">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="REVISEACCT_Name" HeaderText="申請人姓名">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="modifydate" HeaderText="審核時間">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="計畫狀態">
                                            <HeaderStyle Width="10%" />
                                            <ItemTemplate>
                                                <asp:Label ID="labPrjstatus" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center" HeaderStyle-Width="10%" HeaderStyle-HorizontalAlign="Center">
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:LinkButton ID="But_Dir" runat="server" Text="申請變更" CausesValidation="false" CssClass="linkbutton" CommandName="appChg"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table id="tb_PageControler1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <%--<asp:ValidationSummary ID="TotalMsg" runat="server" ShowSummary="False" ShowMessageBox="True" DisplayMode="List"></asp:ValidationSummary>--%>
                </td>
            </tr>
        </table>
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="orgname" type="hidden" name="orgname" runat="server" />
        <input id="ROC_Years" type="hidden" runat="server" />
        <input id="hid_USE_PLAN_REVISESUB" type="hidden" runat="server" />

    </form>
</body>
</html>
