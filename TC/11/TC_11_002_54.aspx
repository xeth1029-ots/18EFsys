<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" EnableEventValidation="false" CodeBehind="TC_11_002_54.aspx.vb" Inherits="WDAIIP.TC_11_002_54" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>線上送件確認</title>
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function selectOnlyThis(id, iMax, DGNN) {
            var s_chkItem1 = "_chkItem1";
            for (var i = 2; i <= iMax + 1; i++) {
                var chkItem1 = null;
                var strID = DGNN + "__ctl" + i + s_chkItem1;
                chkItem1 = document.getElementById(strID);
                if (!chkItem1) {
                    strID = (i < 10) ? DGNN + "_ctl0" + i + s_chkItem1 : DGNN + "_ctl" + i + s_chkItem1;
                    chkItem1 = document.getElementById(strID);
                }
                if (strID != id && chkItem1) { chkItem1.checked = false; }
            }
            //document.getElementById(id).checked = true;
        }
    </script>
    <style type="text/css">
        .Brown-style1 { color: #990000; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;線上送件確認</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTableSch1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <table id="Table3" class="table_sch" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" colspan="3" width="80%">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Button3" type="button" value="..." name="Button3" runat="server" class="button_b_Mini" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <%--<input id="orgid_value" type="hidden" name="orgid_value" runat="server" />--%>
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="50%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">案件編號</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="sch_txtBCASENO" runat="server" Width="40%" MaxLength="30"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">計畫年度 </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="sch_ddlYEARS" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <%--<tr><td class="bluecol_need">申請階段</td><td class="whitecol">
                            <asp:DropDownList ID="sch_ddlAPPSTAGE" runat="server"></asp:DropDownList></td></tr>--%>
                        <tr>
                            <td class="bluecol">申辦人姓名</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="sch_txtBINAME" runat="server" Width="40%" MaxLength="30"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">申辦日期 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="sch_txtBIDATE1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('sch_txtBIDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~<asp:TextBox ID="sch_txtBIDATE2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('sch_txtBIDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">審查狀態</td>
                            <td colspan="3" class="whitecol">
                                <%--'CASE a.BISTATUS WHEN 'B' THEN '已送件' WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過'" & vbCrLf--%>
                                <asp:RadioButtonList ID="rbAPPLIEDRESULT" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="B" Selected="True">已申辦</asp:ListItem>
                                    <asp:ListItem Value="Y">申辦確認</asp:ListItem>
                                    <asp:ListItem Value="R">申辦退件修正</asp:ListItem>
                                    <asp:ListItem Value="N">申辦不通過</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol" colspan="4">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="BTN_SEARCH1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="BTN_TRAIN_PACKAGE_DOWNLOAD1" runat="server" Text="【訓練班別計畫表】分署打包下載" CommandName="TRAIN_PACKAGE_DOWNLOAD1" CssClass="asp_Export_M"></asp:Button>
                                <%--'Training class schedule】Package download--%>
                                <%--<asp:Button ID="BTN_ADDNEW1" runat="server" Text="新增申辦案件" CssClass="asp_button_M"></asp:Button>--%>
                            </td>
                        </tr>
                    </table>
                    <table id="TableDataGrid1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td>
                                <!--案件編號 計畫年度 申請階段 管控單位 機構名稱 申辦人姓名 申辦日期 申辦狀態 功能-->
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" AllowCustomPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="BCASENO" HeaderText="案件編號">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="YEARS_ROC" HeaderText="計畫年度">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="APPSTAGE_N" HeaderText="申請階段"><HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="DISTNAME" HeaderText="管控單位">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="機構名稱">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="BINAME" HeaderText="申辦人姓名">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="BIDDATE_ROC" HeaderText="申辦日期">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="BISTATUS_N" HeaderText="申辦狀態">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="APPLIEDRESULT_N" HeaderText="審查狀態">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lBTN_EDIT1" runat="server" Text="審查" CommandName="EDIT1" CssClass="linkbutton"></asp:LinkButton>&nbsp;
                                                <asp:LinkButton ID="lBTN_REVERT2" runat="server" Text="還原" CommandName="REVERT2" CssClass="linkbutton"></asp:LinkButton>&nbsp;
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
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Label ID="labmsg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTableEdt1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <table id="Table4" class="table_sch" cellpadding="1" cellspacing="1">
                        <tr>
                            <td colspan="4" class="table_title" width="100%">申辦訓練機構資訊</td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" colspan="3" width="80%">
                                <asp:Label ID="labOrgNAME" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">案件編號</td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labBCASENO" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">計畫年度 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labBIYEARS" runat="server"></asp:Label></td>
                            <%--width="30%" td class="bluecol" width="20%">申請階段 </td>
                            <td class="whitecol" width="30%"><asp:Label ID="labAPPSTAGE" runat="server"></asp:Label></td --%>
                        </tr>
                        <tr>
                            <td class="bluecol">班級名稱</td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labCLASSNAME2S" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">檔案下載</td>
                            <td class="whitecol" colspan="3">
                                <%--File package download--%>
                                <asp:Button ID="BTN_PACKAGE_DOWNLOAD1" runat="server" Text="檔案打包下載" CommandName="PACKAGE_DOWNLOAD1" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                        <tr id="tr_HISREVIEW" runat="server">
                            <td class="bluecol">歷程資訊</td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labHISREVIEW" runat="server" Text=""></asp:Label>
                            </td>
                        </tr>
                        <%--<tr><td class="bluecol">切換至</td><td class="whitecol" colspan="3"><asp:DropDownList ID="ddlSwitchTo" runat="server" AppendDataBoundItems="True" AutoPostBack="True"></asp:DropDownList></td></tr>--%>
                        <tr>
                            <td colspan="4" class="table_title" width="100%">線上送件項目</td>
                        </tr>
                        <%--<tr><td class="bluecol" width="20%">線上申辦進度</td><td class="whitecol" colspan="3" width="80%"><asp:Label ID="labProgress" runat="server"></asp:Label></td></tr>--%>
                        <tr>
                            <td align="center" class="whitecol" colspan="4">
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle />
                                    <Columns>
                                        <%--<asp:BoundColumn DataField="KBSID" HeaderText="項目"><HeaderStyle Width="10%" /></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="KBNAME" HeaderText="項目名稱">
                                            <HeaderStyle Width="30%" />
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="BCFID" HeaderText="序號"><HeaderStyle Width="10%" /></asp:BoundColumn>--%>
                                        <%--<asp:BoundColumn DataField="FILENAME1" HeaderText="檔案名稱1"><HeaderStyle Width="10%" /></asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="SRCFILENAME1" HeaderText="上傳檔案">
                                            <HeaderStyle Width="20%" />
                                        </asp:BoundColumn>
                                        <%-- <asp:TemplateColumn HeaderText="檔案名稱"><HeaderStyle Width="20%" /><ItemTemplate>
                                           <asp:Label ID="LabFileName1" runat="server"></asp:Label>
                                           <input id="HFileName" type="hidden" runat="server" /></ItemTemplate></asp:TemplateColumn>--%>
                                        <asp:TemplateColumn HeaderText="退件修正">
                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle" Width="18%"></HeaderStyle>
                                            <ItemTemplate>
                                                <div>
                                                    <span>退回原因說明</span><br />
                                                    <asp:TextBox ID="txtRtuReason" runat="server" MaxLength="300" TextMode="MultiLine" Width="95%" Height="60px"></asp:TextBox><br>
                                                    <asp:Button ID="Btn_RtuBACK1" runat="server" Text="退回開放修改" CommandName="RtuBACK1" CssClass="asp_button_M"></asp:Button>
                                                    <asp:Button ID="Btn_REVERT1" runat="server" Text="還原" CommandName="REVERT1" CssClass="asp_button_M"></asp:Button>
                                                </div>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <%--<asp:Button ID="BTN_DELFILE4" runat="server" Text="查詢" CommandName="DELFILE4" CssClass="asp_button_M"></asp:Button>--%>
                                                <asp:Button ID="BTN_VIEWFILE4" runat="server" Text="查詢" CommandName="VIEWFILE4" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="BTN_DOWNLOAD4" runat="server" Text="下載" CommandName="DOWNLOAD4" CssClass="asp_Export_M"></asp:Button>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <%--<tr><td class="bluecol" width="20%">文件說明</td><td class="whitecol" colspan="3" width="80%">
                                <asp:Literal ID="LiteralSwitchTo" runat="server"></asp:Literal></td></tr><tr><td class="bluecol">檔案格式說明</td>
                                <td class="whitecol" colspan="3">PDF(掃瞄畫面需清楚，檔案大小2MB以下)</td></tr><tr id="tr_DOWNLOADRPT1" runat="server">
                                <td class="bluecol">下載報表 </td><td colspan="3" class="whitecol">
                                <asp:Button ID="BTN_DOWNLOADRPT1" runat="server" Text="下載報表" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td></tr>--%>
                                <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr id="tr_LabSwitchTo" runat="server">
                                        <td colspan="4" class="class_title2_left">
                                            <asp:Label ID="LabSwitchTo" runat="server" Text=""></asp:Label></td>
                                    </tr>

                                    <tr id="tr_DataGrid10" runat="server">
                                        <td colspan="4" class="whitecol">
                                            <%--訓練機構管理>開班資料設定>師資資料設定--%>
                                            <asp:DataGrid ID="DataGrid10" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#f5f5f5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <%--<asp:TemplateColumn HeaderStyle-Width="6%">
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        <ItemTemplate>
                                                            <input type="checkbox" id="chkItem1" data-role="chkItem1" runat="server" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>--%>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="TeachCName" HeaderText="講師名稱">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="IDNO_MK" HeaderText="身分證">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="KINDNAME" HeaderText="師資別">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="KINDENGAGE_N" HeaderText="內外聘">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="SRCFILENAME1" HeaderText="上傳檔案">
                                                        <HeaderStyle Width="10%" />
                                                    </asp:BoundColumn>
                                                    <%-- <asp:TemplateColumn HeaderText="檔案名稱">
                                                        <HeaderStyle Width="10%" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="LabFileName1" runat="server"></asp:Label>
                                                            <input id="HFileName" type="hidden" runat="server" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>--%>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="8%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="HDG10_TechID" type="hidden" runat="server" />
                                                            <input id="HDG10_RID" type="hidden" runat="server" />
                                                            <asp:Button ID="BTN_DOWNLOAD10" runat="server" Text="下載" CommandName="DOWNLOAD10" CssClass="asp_Export_M" />
                                                            <%--<asp:Button ID="BTN_DOWNLOAD10" runat="server" Text="下載" CommandName="DOWNLOAD10" CssClass="asp_button_M" />
                                                            <asp:Button ID="BTN_DELFILE10" runat="server" Text="刪除" CommandName="DELFILE10" CssClass="asp_button_M"></asp:Button>--%>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr id="tr_DataGrid11" runat="server">
                                        <td colspan="4" class="whitecol">
                                            <asp:DataGrid ID="DataGrid11" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#f5f5f5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <%--<asp:TemplateColumn HeaderStyle-Width="6%">
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        <ItemTemplate>
                                                            <input type="checkbox" id="chkItem1" data-role="chkItem1" runat="server" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>--%>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="TeachCName" HeaderText="講師名稱">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="IDNO_MK" HeaderText="身分證">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="KINDNAME" HeaderText="師資別">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="KINDENGAGE_N" HeaderText="內外聘">
                                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="SRCFILENAME1" HeaderText="上傳檔案">
                                                        <HeaderStyle Width="10%" />
                                                    </asp:BoundColumn>
                                                    <%--<asp:TemplateColumn HeaderText="檔案名稱">
                                                        <HeaderStyle Width="10%" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="LabFileName1" runat="server"></asp:Label>
                                                            <input id="HFileName" type="hidden" runat="server" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>--%>
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="8%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="HDG11_TechID" type="hidden" runat="server" />
                                                            <input id="HDG11_RID" type="hidden" runat="server" />
                                                            <asp:Button ID="BTN_DOWNLOAD11" runat="server" Text="下載" CommandName="DOWNLOAD11" CssClass="asp_Export_M" />
                                                            <%--<asp:Button ID="BTN_DELFILE11" runat="server" Text="刪除" CommandName="DELFILE11" CssClass="asp_button_M"></asp:Button>--%>
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr id="tr_DataGrid13" runat="server">
                                        <td colspan="4" class="whitecol">
                                            <asp:DataGrid ID="DataGrid13" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#f5f5f5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="CLASSCNAMEX" HeaderText="班級名稱">
                                                        <HeaderStyle Width="60%" />
                                                    </asp:BoundColumn>
                                                    <%--  <asp:TemplateColumn HeaderText="檔案名稱">
                                                        <HeaderStyle Width="10%" />
                                                        <ItemTemplate>
                                                            <asp:Label ID="LabFileName1" runat="server"></asp:Label>
                                                            <input id="HFileName" type="hidden" runat="server" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>--%>
                                                    <asp:TemplateColumn>
                                                        <HeaderStyle Width="11%" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="HDG_PlanID" type="hidden" runat="server" />
                                                            <input id="HDG_ComIDNO" type="hidden" runat="server" />
                                                            <input id="HDG_SeqNo" type="hidden" runat="server" />
                                                            <asp:Button ID="BTN_DOWNLOAD13" runat="server" Text="下載" CommandName="DOWNLOAD13" CssClass="asp_Export_M" />
                                                            <asp:Button ID="BTN_REPORT13" runat="server" Text="列印" CommandName="REPORT13" CssClass="asp_Export_M" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr id="tr_DataGrid13B" runat="server">
                                        <td colspan="4" class="whitecol">
                                            <asp:DataGrid ID="DataGrid13B" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#f5f5f5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="CLASSCNAMEX" HeaderText="班級名稱">
                                                        <HeaderStyle Width="60%" />
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn>
                                                        <HeaderStyle Width="11%" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="HDG_PlanID" type="hidden" runat="server" />
                                                            <input id="HDG_ComIDNO" type="hidden" runat="server" />
                                                            <input id="HDG_SeqNo" type="hidden" runat="server" />
                                                            <asp:Button ID="BTN_DOWNLOAD13B" runat="server" Text="下載" CommandName="DOWNLOAD13B" CssClass="asp_Export_M" />
                                                            <asp:Button ID="BTN_REPORT13B" runat="server" Text="列印" CommandName="REPORT13B" CssClass="asp_Export_M" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr id="tr_DataGrid14" runat="server">
                                        <td colspan="4" class="whitecol">
                                            <asp:DataGrid ID="DataGrid14" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#f5f5f5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <asp:BoundColumn HeaderText="序號">
                                                        <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="CLASSCNAMEX" HeaderText="班級名稱">
                                                        <HeaderStyle Width="50%" />
                                                    </asp:BoundColumn>
                                                    <asp:BoundColumn DataField="SRCFILENAME1" HeaderText="上傳檔案">
                                                        <HeaderStyle Width="11%" />
                                                    </asp:BoundColumn>
                                                    <asp:TemplateColumn>
                                                        <HeaderStyle Width="11%" />
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <input id="HDG_PlanID" type="hidden" runat="server" />
                                                            <input id="HDG_ComIDNO" type="hidden" runat="server" />
                                                            <input id="HDG_SeqNo" type="hidden" runat="server" />
                                                            <asp:Button ID="BTN_DOWNLOAD14" runat="server" Text="下載" CommandName="DOWNLOAD14" CssClass="asp_Export_M" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                </table>
                                <%--
<tr id="tr_SENTBATVER" runat="server"><td class="bluecol">以目前版本批次送出</td><td colspan="3" class="whitecol"><asp:Button ID="BTN_SENTBATVER" runat="server" Text="以目前版本批次送出" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td></tr>
<tr id="tr_SENDCURRVER" runat="server"><td class="bluecol">以目前版本送出</td><td colspan="3" class="whitecol"><asp:Button ID="BTN_SENDCURRVER" runat="server" Text="以目前版本送出" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td></tr>
<tr id="tr_UPLOADFL1" runat="server"><td class="bluecol">檔案上傳 </td><td colspan="3" class="whitecol"><input id="File1" type="file" size="66" name="File1" runat="server" accept=".pdf" /><asp:Button ID="But1" runat="server" Text="確定檔案上傳" CausesValidation="False" CssClass="asp_button_M"></asp:Button></td></tr>
<tr id="tr_USELATESTVER" runat="server"><td class="bluecol">最近一次版本送件</td><td class="whitecol" colspan="3"><asp:Button ID="bt_latestSend1" runat="server" Text="最近一次版本送件" /></td></tr>
<tr id="tr_WAIVED" runat="server"><td class="bluecol">免附文件</td><td class="whitecol" colspan="3"><asp:CheckBox ID="CHKB_WAIVED" runat="server" />免附文件</td></tr>
<tr id="tr_USEMEMO1" runat="server"><td class="bluecol">備註說明</td><td class="whitecol" colspan="3"><asp:TextBox ID="txtMEMO1" runat="server" Width="60%"></asp:TextBox></td></tr>
                                --%>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="table_title" width="100%">資料審查結果</td>
                        </tr>
                        <tr>
                            <td class="bluecol">審查狀態 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddlAPPLIEDRESULT" runat="server">
                                    <asp:ListItem Value="">===請選擇===</asp:ListItem>
                                    <asp:ListItem Value="Y">申辦確認</asp:ListItem>
                                    <asp:ListItem Value="R">申辦退件修正</asp:ListItem>
                                    <asp:ListItem Value="N">申辦不通過</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">不通過原因 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="Reasonforfail" runat="server" TextMode="MultiLine" Rows="7" Width="88%" MaxLength="500"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="whitecol" align="center">
                                <asp:Button ID="But_Sub" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="But_BACK1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Label ID="labmsg2" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>

        <%--目前項次(流水號)Hid_KBSID／目前項次(序號)Hid_KBID--%>
        <asp:HiddenField ID="Hid_KBSID" runat="server" />
        <asp:HiddenField ID="Hid_KBID" runat="server" />
        <asp:HiddenField ID="Hid_TECHID" runat="server" />

        <%--申辦流水號：Hid_BCID／CASENO：Hid_BCASENO／ORGKINDGW：Hid_ORGKINDGW／申請班級號pcs(多筆)：Hid_PCS--%>
        <asp:HiddenField ID="Hid_BCID" runat="server" />
        <asp:HiddenField ID="Hid_RID" runat="server" />
        <asp:HiddenField ID="Hid_BCFID" runat="server" />
        <asp:HiddenField ID="Hid_BCASENO" runat="server" />
        <asp:HiddenField ID="Hid_ORGKINDGW" runat="server" />
        <asp:HiddenField ID="Hid_PCS" runat="server" />
        <asp:HiddenField ID="Hid_APPLIEDRESULT" runat="server" />
        <asp:HiddenField ID="Hid_USE_ORG_BIDCASE_REVERT2" runat="server" />
        <%--<asp:HiddenField ID="HiddenField1" runat="server" />--%>
    </form>
</body>
</html>
