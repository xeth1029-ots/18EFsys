<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_04_004_17.aspx.vb" Inherits="WDAIIP.TC_04_004_17" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>重點產業審核確認</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;重點產業審核確認</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Panel ID="panelSearch" runat="server">
                        <table id="Table3" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" width="20%">訓練機構 </td>
                                <td colspan="3" class="whitecol">
                                    <asp:TextBox ID="center" runat="server" Width="40%" Myonfocus="this.blur()"></asp:TextBox><input id="Org" value="..." type="button" name="Org" runat="server" class="asp_button_M">
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                    <span style="position: absolute; display: none" id="HistoryList2">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span></td>
                            </tr>
                            <tr>
                                <td class="bluecol">
                                    <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label>
                                </td>
                                <td colspan="3" class="whitecol">
                                    <asp:TextBox ID="TB_career_id" runat="server" Myonfocus="this.blur()" Columns="40"></asp:TextBox>
                                    <input id="btu_sel" onclick="openTrain(document.getElementById('trainValue').value);" value="..." type="button" name="btu_sel" runat="server" class="asp_button_M">
                                    <input id="TPlanid" type="hidden" name="TPlanid" runat="server">
                                    <input id="trainValue" type="hidden" name="trainValue" runat="server">
                                    <input id="jobValue" type="hidden" name="jobValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">
                                    <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                                </td>
                                <td colspan="3" class="whitecol">
                                    <asp:TextBox ID="txtCJOB_NAME" runat="server" Myonfocus="this.blur()" Columns="40"></asp:TextBox><input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" value="..." type="button" name="btu_sel2" runat="server" class="asp_button_M">
                                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">班別名稱 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="ClassName" runat="server" Columns="40" MaxLength="50"></asp:TextBox>
                                </td>
                                <td class="bluecol" width="20%">期別 </td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="CyclType" runat="server" Columns="5" MaxLength="2"></asp:TextBox>
                                </td>
                            </tr>
                            <tr id="tr_txtCLASSNAMEKeyWORDS_SCH" runat="server">
                                <td class="bluecol">課程關鍵字 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="txtCLASSNAMEKeyWORDS_SCH" runat="server" Columns="40" MaxLength="50"></asp:TextBox>
                                </td>
                            </tr>
                            <%-- <tr>
                                <td class="bluecol">課程申請流水號 </td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="txtPSNO28_SCH" runat="server" Columns="40" MaxLength="20"></asp:TextBox>
                                </td>
                            </tr>--%>
                            <tr>
                                <td class="bluecol">開訓日期 </td>
                                <td colspan="3" class="whitecol">
                                    <span runat="server">
                                        <asp:TextBox ID="STDate1" Width="15%" Myonfocus="this.blur()" runat="server"></asp:TextBox>
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate1.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif">～
								        <asp:TextBox ID="STDate2" Width="15%" Myonfocus="this.blur()" runat="server"></asp:TextBox>
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate2.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif">
                                    </span></td>
                            </tr>
                            <tr>
                                <td class="bluecol">結訓日期 </td>
                                <td colspan="3" class="whitecol">
                                    <span runat="server">
                                        <asp:TextBox ID="FDDate1" Width="15%" Myonfocus="this.blur()" runat="server"></asp:TextBox>
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FDDate1.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif">～
								        <asp:TextBox ID="FDDate2" Width="15%" Myonfocus="this.blur()" runat="server"></asp:TextBox>
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FDDate2.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif">
                                    </span></td>
                            </tr>
                            <tr>
                                <td class="bluecol">計畫範圍 </td>
                                <td colspan="3" class="whitecol">
                                    <asp:RadioButtonList ID="OrgKind2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                        <asp:ListItem Value="G">產業人才投資計畫</asp:ListItem>
                                        <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr id="tr_AppStage_TP28" runat="server">
                                <td class="bluecol">申請階段</td>
                                <td class="whitecol" colspan="3">
                                    <asp:RadioButtonList ID="AppStage2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font"></asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級審核類型 </td>
                                <td colspan="3" class="whitecol">
                                    <asp:RadioButtonList ID="PlanMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="S" Selected="True">審核中</asp:ListItem>
                                        <asp:ListItem Value="Y">已通過</asp:ListItem>
                                        <asp:ListItem Value="R">退件修正(含不通過的)</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr id="tr_Btn_XlsImport1" runat="server">
                                <td class="bluecol">產業別分類匯入 </td>
                                <td colspan="3" class="whitecol">
                                    <div align="left">
                                        <input id="File1" type="file" size="80" name="File1" runat="server" accept=".xls,.ods" />
                                        <asp:Button ID="Btn_XlsImport1" runat="server" Text="匯入產業別分類" CssClass="asp_Export_M"></asp:Button>(必須為ods或xls格式)<asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" ForeColor="#8080FF">下載 (匯入產業別分類)格式檔</asp:HyperLink><%--&nbsp;&nbsp;--%>
                                    </div>
                                </td>
                            </tr>
                            <tr id="tr_RBListExpType" runat="server">
                                <td class="bluecol">匯出檔案格式</td>
                                <td colspan="3" class="whitecol">
                                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" align="center" class="whitecol">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnExp1" runat="server" Text="匯出產業別分類" CssClass="asp_Export_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table id="DataGridTable1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td align="center">
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <ItemStyle></ItemStyle>
                                        <AlternatingItemStyle BackColor="WhiteSmoke" />
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <SelectedItemStyle></SelectedItemStyle>
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="序號">
                                                <HeaderStyle Width="3%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="Years" HeaderText="年度">
                                                <HeaderStyle Width="7%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="DistName" HeaderText="轄區">
                                                <HeaderStyle Width="8%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                                <HeaderStyle Width="9%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱">
                                                <HeaderStyle Width="9%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <%--,ig3.GCODE31,ig3.PNAME JobName (課程分類／職類課程)--%>
                                            <%--<asp:BoundColumn DataField="PNAME" HeaderText="職類課程"><HeaderStyle Width="9%"></HeaderStyle></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="JobName" HeaderText="課程分類">
                                                <HeaderStyle Width="8%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <%--<asp:TemplateColumn HeaderText="課程分類"><HeaderStyle Width="8%"></HeaderStyle><ItemTemplate><asp:Label ID="lbD12KNAME" runat="server"></asp:Label></ItemTemplate></asp:TemplateColumn>--%>
                                            <asp:BoundColumn DataField="CNAME" HeaderText="訓練業別">
                                                <HeaderStyle Width="8%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="JobName " HeaderText="訓練業別"><HeaderStyle Width="8%"></HeaderStyle></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="GovClassN" HeaderText="訓練業別編碼">
                                                <HeaderStyle Width="7%"></HeaderStyle>
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="轄區重點產業">
                                                <HeaderStyle Width="7%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lbD15KNAME" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <%--<asp:TemplateColumn HeaderText="新興產業"><HeaderStyle Width="7%"></HeaderStyle><ItemTemplate><asp:Label ID="lbD06KNAME" runat="server"></asp:Label></ItemTemplate></asp:TemplateColumn>
                                          <asp:TemplateColumn HeaderText="重點服務業"><HeaderStyle Width="7%"></HeaderStyle><ItemTemplate><asp:Label ID="lbD10KNAME" runat="server"></asp:Label></ItemTemplate></asp:TemplateColumn>--%>
                                            <asp:TemplateColumn HeaderText="政府政策性產業">
                                                <HeaderStyle Width="8%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <%--<asp:Label ID="lbD17KNAME" runat="server"></asp:Label><asp:Label ID="lbD19KNAME" runat="server"></asp:Label>--%>
                                                    <asp:Label ID="lbD20KNAME" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <%-- <asp:TemplateColumn HeaderText="新南向政策"><HeaderStyle Width="7%"></HeaderStyle><ItemTemplate><asp:Label ID="lbD18KNAME" runat="server"></asp:Label></ItemTemplate></asp:TemplateColumn>--%>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <HeaderStyle Width="4%" />
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <HeaderTemplate>
                                                    <asp:Label ID="lltitle" runat="server">功能</asp:Label>
                                                </HeaderTemplate>
                                                <ItemTemplate>
                                                    <asp:Button ID="BtnCHKOK" runat="server" Text="確認" CommandName="CHKOK" CssClass="asp_button_M"></asp:Button>
                                                    <asp:Button ID="BtnEdit" runat="server" Text="修改" CommandName="Edit" CssClass="asp_button_M"></asp:Button>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td style="height: 31px" align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="PanelEdit1" runat="server">
                        <table id="Table4" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" width="20%">年度 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbYears" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">轄區 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbDistName" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練機構 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbOrgName" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級名稱 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbClassName" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <%--<tr>
                                <td class="bluecol">課程申請流水號／班級代碼 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbPSNO28" runat="server"></asp:Label>
                                    ／<asp:Label Style="z-index: 0" ID="lbOCID" runat="server"></asp:Label></td>
                            </tr>--%>
                            <tr>
                                <td class="bluecol">班級代碼 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbOCID" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練期間 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbSFTDate" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練時數 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbTHours" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練業別 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbJobName" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練業別編碼 </td>
                                <td>
                                    <asp:Label Style="z-index: 0" ID="lbGovClassN" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td class="table_title" colspan="2">課程大綱 </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <table id="Datagrid3Table" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="Datagrid3" runat="server" CssClass="font" AutoGenerateColumns="False" Width="100%" CellPadding="8">
                                                    <EditItemStyle Wrap="False"></EditItemStyle>
                                                    <ItemStyle></ItemStyle>
                                                    <AlternatingItemStyle BackColor="WhiteSmoke" />
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="時數">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="PHourLabel" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="課程進度／內容">
                                                            <HeaderStyle Width="80%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="lbContText" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="學／術科">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" CssClass="whitecol"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:DropDownList ID="drpClassification1" runat="server" Enabled="False" AutoPostBack="True">
                                                                    <asp:ListItem Value="1">學科</asp:ListItem>
                                                                    <asp:ListItem Value="2">術科</asp:ListItem>
                                                                </asp:DropDownList>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                    <PagerStyle Visible="False"></PagerStyle>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">課程分類 </td>
                                <td class="whitecol">
                                    <asp:DropDownList Style="z-index: 0" ID="ddlDEPOT12" runat="server">
                                    </asp:DropDownList>
                                    <span style="color: #FF0000;">(除點選「其他類」，餘應修正「經費分類代碼」)</span> </td>
                            </tr>
                            <tr>
                                <td class="bluecol">轄區重點產業 </td>
                                <td class="whitecol">
                                    <asp:DropDownList Style="z-index: 0" ID="ddlDEPOT15" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <%--<tr><td class="bluecol">新興產業 </td><td class="whitecol"><asp:DropDownList Style="z-index: 0" ID="ddlKID06" runat="server"></asp:DropDownList></td></tr>
                            <tr><td class="bluecol">重點服務業 </td><td class="whitecol"><asp:DropDownList Style="z-index: 0" ID="ddlKID10" runat="server"></asp:DropDownList></td></tr>
                                <tr id="trKID19" runat="server"><td class="bluecol">政府政策性產業 </td><td class="whitecol"><asp:DropDownList Style="z-index: 0" ID="ddlKID17" runat="server"></asp:DropDownList>
                                <asp:DropDownList Style="z-index: 0" ID="ddlKID19" runat="server"></asp:DropDownList></td></tr><tr id="trKID18" runat="server"><td class="bluecol">新南向政策 </td><td class="whitecol">
                                <asp:DropDownList Style="z-index: 0" ID="ddlKID18" runat="server"></asp:DropDownList></td></tr>--%>
                            <tr id="trKID20" runat="server">
                                <td class="bluecol" width="20%">&nbsp; 政府政策性產業 </td>
                                <td class="whitecol" width="80%">
                                    <table border="0" cellspacing="1" cellpadding="1" width="100%">
                                        <tr>
                                            <td class="bluecol" width="22%">「5+2」產業創新計畫</td>
                                            <td class="whitecol" width="78%">
                                                <asp:CheckBoxList ID="CBLKID20_1" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">台灣AI行動計畫</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID20_2" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">數位國家創新經濟<br />
                                                發展方案</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID20_3" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">國家資通安全發展方案</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID20_4" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">前瞻基礎建設計畫</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID20_5" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">新南向政策</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID20_6" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">進階政策性產業類別</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID22" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="trKID25" runat="server">
                                <td class="bluecol">&nbsp; 政府政策性產業 </td>
                                <td class="whitecol">
                                    <table border="0" cellspacing="1" cellpadding="1" width="100%">
                                        <tr>
                                            <td class="bluecol" width="22%">亞洲矽谷</td>
                                            <td class="whitecol" width="78%">
                                                <asp:CheckBoxList ID="CBLKID25_1" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">重點產業</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_2" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">台灣AI行動計畫</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_3" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">智慧國家方案</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_4" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">國家人才競爭力躍升方案</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_5" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">新南向政策</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_6" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr id="trCBLKID25_7" runat="server">
                                            <td class="bluecol">AI加值應用</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_7" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr id="trCBLKID25_8" runat="server">
                                            <td class="bluecol">職場續航</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID25_8" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">進階政策性產業類別</td>
                                            <td class="whitecol">
                                                <asp:CheckBoxList ID="CBLKID22B" runat="server" RepeatColumns="3" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="tr_CBLKID60" runat="server">
                                <td class="bluecol">產業別(管考) </td>
                                <td class="whitecol">
                                    <asp:CheckBoxList ID="CBLKID60" runat="server" RepeatColumns="4" RepeatDirection="Horizontal"></asp:CheckBoxList>
                                </td>
                            </tr>
                            <%--<tr> <td class="bluecol" width="100"> &nbsp; 新興智慧型產業 </td>
							<td> <asp:DropDownList Style="z-index: 0" ID="ddlKID04" runat="server"></asp:DropDownList></td></tr>--%>
                            <tr>
                                <td colspan="2" align="center" class="whitecol">
                                    <asp:Button ID="btnSave1" runat="server" Text="確認" CssClass="asp_button_M"></asp:Button>&nbsp;
								<asp:Button ID="btnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
            <%--<tr><td class="bluecol" width="100">&nbsp; 生產力4.0</td><td>	
                <asp:DropDownList Style="z-index: 0" ID="ddlKID14" runat="server"></asp:DropDownList></td></tr>--%>
        </table>
        <input id="hidPlanID" type="hidden" name="hidPlanID" runat="server">
        <input id="hidComIDNO" type="hidden" name="hidComIDNO" runat="server">
        <input id="hidSeqNO" type="hidden" name="hidSeqNO" runat="server">
        <asp:HiddenField ID="Hid_Years2018" runat="server" />
        <asp:HiddenField ID="Hid_Years2019" runat="server" />
        <asp:HiddenField ID="Hid_TMIDVALUE" runat="server" />
    </form>
</body>
</html>
