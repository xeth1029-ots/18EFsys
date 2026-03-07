<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" EnableEventValidation="false" CodeBehind="TC_13_002.aspx.vb" Inherits="WDAIIP.TC_13_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>核銷文件審查</title>
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;補助金請領&gt;&gt;核銷文件審查</asp:Label>
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
                            <td class="bluecol_need">職類/班別 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <span id="HistoryList" style="display: none; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">申辦人姓名</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="sch_txtSENDACCTNAME" runat="server" Width="40%" MaxLength="30"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">申辦日期 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="sch_txtSENDDATE1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('sch_txtSENDDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~<asp:TextBox ID="sch_txtSENDDATE2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('sch_txtSENDDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">審查狀態</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="rbAPPLIEDRESULT" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="B" Selected="True">已申辦</asp:ListItem>
                                    <asp:ListItem Value="Y">申辦確認</asp:ListItem>
                                    <asp:ListItem Value="R">申辦退件修正</asp:ListItem>
                                    <asp:ListItem Value="N">申辦不通過</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol" colspan="4">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="BTN_SEARCH1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <%--<asp:Button ID="BTN_TRAIN_PACKAGE_DOWNLOAD1" runat="server" Text="【訓練班別計畫表】分署打包下載" CommandName="TRAIN_PACKAGE_DOWNLOAD1" CssClass="asp_Export_M"></asp:Button>--%>
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
                                        <asp:BoundColumn DataField="YEARS_ROC" HeaderText="計畫年度">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="APPSTAGE_N" HeaderText="申請階段">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="DISTNAME" HeaderText="管控單位">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SENDDATE_ROC" HeaderText="申辦日期">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SENDSTATUS_N" HeaderText="申辦狀態">
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
                        <%-- <tr>,<td class="bluecol">案件編號</td>,<td class="whitecol" colspan="3">,<asp:Label ID="labBCASENO" runat="server"></asp:Label></td>,</tr>--%>
                        <tr>
                            <td class="bluecol" width="20%">計畫年度 </td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="labSEND_YEARS_ROC" runat="server"></asp:Label></td>
                            <td class="bluecol" width="20%">申請階段 </td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="labAPPSTAGE" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級名稱</td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labCLASSNAME2" runat="server"></asp:Label></td>
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
                                        <%-- <asp:TemplateColumn HeaderText="檔案名稱"><HeaderStyle Width="20%" /><ItemTemplate>,<asp:Label ID="LabFileName1" runat="server"></asp:Label>,
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
                        <%-- <tr><td colspan="4"></td></tr>--%>
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
        <asp:HiddenField ID="Hid_ORGKINDGW" runat="server" />
        <%-- 目前項次(流水號)Hid_KVSID／目前項次(序號)Hid_KVID--%>
        <asp:HiddenField ID="Hid_KVSID" runat="server" />
        <asp:HiddenField ID="Hid_KVID" runat="server" />
        <asp:HiddenField ID="Hid_LastKVID" runat="server" />
        <asp:HiddenField ID="Hid_FirstKVSID" runat="server" />
        <asp:HiddenField ID="Hid_CVOCFID" runat="server" />
        <%--申辦--%>
        <asp:HiddenField ID="Hid_CVOCID" runat="server" />
        <asp:HiddenField ID="Hid_OCIDVal" runat="server" />
        <asp:HiddenField ID="Hid_SEQ_ID" runat="server" />
        <asp:HiddenField ID="Hid_APPLIEDRESULT" runat="server" />
        <%--<asp:HiddenField ID="HiddenField1" runat="server" />--%>
    </form>
</body>
</html>
