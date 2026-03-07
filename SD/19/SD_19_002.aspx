<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" EnableEventValidation="true" CodeBehind="SD_19_002.aspx.vb" Inherits="WDAIIP.SD_19_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>開訓資料線上送件確認</title>
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
    <script type="text/javascript">
        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;開訓資料線上送件確認</asp:Label>
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
                            <td class="bluecol">職類/班別</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="Historytable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <%--A29ACD67218B121F132FF90ED282A0AB0CB054585FA86AE05029CD7B61A4CCD882CF7DDD609BA63A2F1D44C9080238B60295090A13BBA9B223372312
B247A333BCBBE7DE1F1F9AA85620B3ADBDE50A13961D16A340F729BBCD38297A6FECCFA53D735009AAD3F8294D7C81B350FA3BA820A8DB1523C6EA84
42575D32374A719B166506DA41365C937D3C188BE37025C3E6B412402C4365E39194DF9BAA8CEA412C495E607ED6B0A148E899300FE4557B--%>
                        <tr>
                            <td class="bluecol_need">計畫年度 </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="sch_ddlYEARS" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">申請階段</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="sch_ddlAPPSTAGE" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol">申辦人姓名</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="sch_txtTBCNAME" runat="server" Width="40%" MaxLength="30"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">申辦日期 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="sch_txtTBCDATE1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('sch_txtTBCDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~<asp:TextBox ID="sch_txtTBCDATE2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('sch_txtTBCDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
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
                                <%--<asp:Button ID="BTN_TRAIN_PACKAGE_DOWNLOAD1" runat="server" Text="分署打包下載" CommandName="TRAIN_PACKAGE_DOWNLOAD1" CssClass="asp_Export_M"></asp:Button>--%>
                                <%--'Training class schedule】Package download--%>
                                <%--<asp:Button ID="BTN_ADDNEW1" runat="server" Text="新增申辦案件" CssClass="asp_button_M"></asp:Button>--%>
                            </td>
                        </tr>
                    </table>
                    <table id="TB_DataGrid1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
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
                                        <asp:BoundColumn DataField="TBCASENO" HeaderText="案件編號">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
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
                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TBCNAME" HeaderText="申辦人姓名">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TBCDATE_ROC" HeaderText="申辦日期">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TBCSTATUS_N" HeaderText="申辦狀態">
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
                            <td colspan="4" class="table_title" width="100%">申辦訓練資訊</td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" colspan="3" width="80%">
                                <asp:Label ID="labOrgNAME" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">案件編號</td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="labTBCASENO" runat="server"></asp:Label></td>
                            <td class="bluecol" width="20%">案件建立時間</td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="LabCREATEDATE" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">計畫年度 </td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="labBIYEARS" runat="server"></asp:Label></td>
                            <td class="bluecol" width="20%">申請階段 </td>
                            <td class="whitecol" width="30%">
                                <asp:Label ID="labAPPSTAGE" runat="server"></asp:Label></td>
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
                                        <asp:BoundColumn DataField="KTNAME2" HeaderText="項目名稱">
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
                            <td id="td_title06" runat="server" colspan="4" class="table_title" width="100%">06.其他補充資料</td>
                        </tr>
                        <tr id="tr_DataGrid06" runat="server">
                            <td colspan="4" class="whitecol">
                                <asp:DataGrid ID="DataGrid06" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="6">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="ROWNUM1" HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SRCFILENAME1" HeaderText="上傳檔案">
                                            <HeaderStyle Width="21%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="MEMO1" HeaderText="備註說明">
                                            <HeaderStyle Width="21%" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="11%" />
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:HiddenField ID="Hid_CS14OFID" runat="server" />
                                                <asp:Button ID="BTN_DOWNLOAD06" runat="server" Text="檔案下載" CommandName="DOWNLOAD06" CssClass="asp_Export_M" />
                                                <%--<asp:Button ID="BTN_DELFILE06" runat="server" Text="刪除檔案" CommandName="DELFILE06" CssClass="asp_button_M" />--%>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
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
        <%--案件索引--%>
        <asp:HiddenField ID="Hid_RID" runat="server" />
        <asp:HiddenField ID="Hid_APPLIEDRESULT" runat="server" />
        <%--<asp:HiddenField ID="Hid_KTSEQ" runat="server" />--%>
        <asp:HiddenField ID="Hid_TBCID" runat="server" />
        <asp:HiddenField ID="Hid_TBCASENO" runat="server" />
        <asp:HiddenField ID="Hid_ORGKINDGW" runat="server" />
        <%--系統參數--%>
        <asp:HiddenField ID="USE_CLASS_STD14OA_REVERT2" runat="server" />
    </form>
</body>
</html>
