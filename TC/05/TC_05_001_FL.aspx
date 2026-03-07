<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" EnableEventValidation="false" CodeBehind="TC_05_001_FL.aspx.vb" Inherits="WDAIIP.TC_05_001_FL" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>班級變更申請</title>
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
        /*
        function ShowFrame() {
            var FrameObj = document.getElementById('FrameObj');
            var HistoryList2 = document.getElementById('HistoryList2');
            FrameObj.style.display = HistoryList2.style.display;
        }
        */
        function checkFile1(sizeLimit) {
            //sizeLimit單位:byte   
            const fileInput = document.querySelector('input[type="file"]');
            const file = fileInput.files[0];
            const fileType = file.type;
            const fileType2 = file.type.split('/')[1];
            //console.log('fileType : ' + fileType);
            //console.log('fileType2 : ' + fileType2);
            if (fileType !== 'application/pdf' || fileType2 !== 'pdf') {
                alert('只允許上傳 PDF 檔！');
                return false;
            }
            const fileSize = Math.round(file.size / 1024 / 1024);
            const fileSizeLimit = Math.round(sizeLimit / 1024 / 1024);
            //console.log('fileSize : ' + fileSize);
            //console.log('fileSizeLimit : ' + fileSizeLimit);
            if (fileSize > fileSizeLimit) {
                alert('您所選擇的檔案大小為 ' + fileSize + 'MB，超過了上傳上限! (檔案大小限制' + fileSizeLimit + 'MB以下)\n不允許上傳！');
                document.getElementById("File1").outerHTML = '<input name="File1" type="file" id="File1" size="66" accept=".pdf">';
                return false;
            }
            return true;
        }
    </script>
    <%--<style type="text/css">
        .Brown-style1 { color: #990000; }
        .auto-style1 { display: inline-block; padding: 6px 16px; border-radius: 2px; background-color: #0eabd6; color: #FFF; margin-left: 2px; margin-right: 2px; margin-bottom: 2px; }
    </style>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;班級變更申請</asp:Label>
                </td>
            </tr>
        </table>

        <table id="FrameTableEdt1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <table id="Table4" class="table_sch" cellpadding="1" cellspacing="1">
                        <tr>
                            <td colspan="4" class="table_title" width="100%">
                                <asp:Label ID="labtitle1" runat="server" Text="班級變更資訊"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="16%">年度 </td>
                            <td class="whitecol" width="34%">
                                <asp:Label ID="YearList" runat="server"></asp:Label><asp:Label ID="labAPPSTAGE" runat="server"></asp:Label></td>
                            <td class="bluecol" width="16%">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td class="whitecol" width="34%">
                                <asp:Label ID="TrainText" runat="server"></asp:Label><asp:Label ID="JobText" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="Labcjob" runat="server">通俗職類</asp:Label></td>
                            <td class="whitecol">
                                <asp:Label ID="CjobName" runat="server"></asp:Label></td>
                            <td class="bluecol">申請人姓名</td>
                            <td class="whitecol">
                                <asp:Label ID="lab_REVISEACCT_Name" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構 </td>
                            <td class="whitecol">
                                <asp:Label ID="OrgName" runat="server"></asp:Label>
                                <input id="RIDValue" type="hidden" runat="server" />
                            </td>
                            <td class="bluecol">班別名稱 </td>
                            <td class="whitecol">
                                <asp:Label ID="ClassName" runat="server"></asp:Label>
                                <asp:Label ID="PointYN" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練期間 </td>
                            <td class="whitecol">
                                <asp:Label ID="TRange" runat="server"></asp:Label></td>
                            <td class="bluecol">是否轉班 </td>
                            <td class="whitecol">
                                <asp:Label ID="ClassFlag" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">查詢模式 </td>
                            <td class="whitecol">
                                <asp:Label ID="SearchMode" runat="server"></asp:Label></td>
                            <td class="bluecol">
                                <asp:Label ID="labTitle" runat="server"></asp:Label></td>
                            <td class="whitecol">
                                <asp:Label ID="CheckMode" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">變更項目 </td>
                            <td class="whitecol">
                                <asp:Label ID="LabChgItem_N" runat="server"></asp:Label>
                                <input id="chgState" type="hidden" value="0" name="chgState" runat="server" />
                            </td>
                            <td class="bluecol_need">申請變更日 </td>
                            <td class="whitecol">
                                <asp:Label ID="labApplyDate_AD" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">切換至</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddlSwitchTo" runat="server" AppendDataBoundItems="True" AutoPostBack="True"></asp:DropDownList>
                                <asp:Button ID="BTN_SEARCH2" runat="server" Text="重新查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="table_title" width="100%">
                                <asp:Label ID="labtitle2" runat="server" Text="線上送件進度"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">線上送件進度</td>
                            <td class="whitecol" colspan="3" width="80%">
                                <asp:Label ID="labProgress" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol" colspan="4">
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <AlternatingItemStyle />
                                    <Columns>
                                        <asp:BoundColumn DataField="RVNAME2" HeaderText="項目名稱">
                                            <HeaderStyle Width="30%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SRCFILENAME1" HeaderText="上傳檔案">
                                            <HeaderStyle Width="22%" />
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle Width="22%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Button ID="BTN_DELFILE4" runat="server" Text="刪除檔案" CommandName="DELFILE4" CssClass="asp_Export_M" />
                                                <asp:Button ID="BTN_DOWNLOAD4" runat="server" Text="檔案下載" CommandName="DOWNLOAD4" CssClass="asp_Export_M" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" class="table_title" width="100%">應備文件上傳</td>
                        </tr>
                        <tr>
                            <td colspan="4" class="class_title2_left">
                                <asp:Label ID="LabSwitchTo" runat="server" Text=""></asp:Label></td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr id="tr_LiteralSwitchTo" runat="server">
                                        <td class="bluecol" width="20%">文件說明</td>
                                        <td class="whitecol" colspan="3" width="80%">
                                            <asp:Literal ID="LiteralSwitchTo" runat="server"></asp:Literal>
                                        </td>
                                    </tr>
                                    <tr id="tr_FILEDESC1" runat="server">
                                        <td class="bluecol">檔案格式說明</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Label ID="labFILEDESC1" runat="server" Text="PDF(掃瞄畫面需清楚，檔案大小10MB以下)"></asp:Label></td>
                                    </tr>
                                    <tr id="tr_DOWNLOADRPT1" runat="server">
                                        <td class="bluecol">下載報表 </td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Button ID="BTN_DOWNLOADRPT1" runat="server" Text="下載報表" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr id="tr_DataGrid10" runat="server">
                                        <td colspan="4" class="whitecol">
                                            <%--訓練機構管理>開班資料設定>師資資料設定--%>
                                            <asp:DataGrid ID="DataGrid10" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                <AlternatingItemStyle BackColor="#f5f5f5" />
                                                <HeaderStyle CssClass="head_navy" />
                                                <Columns>
                                                    <asp:TemplateColumn HeaderStyle-Width="6%">
                                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                        <ItemTemplate>
                                                            <input type="checkbox" id="chkItem1" data-role="chkItem1" runat="server" />
                                                            <input id="HDG10_TechID" type="hidden" runat="server" />
                                                            <input id="HDG10_RID" type="hidden" runat="server" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
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
                                                    <asp:TemplateColumn HeaderText="功能">
                                                        <HeaderStyle Width="8%"></HeaderStyle>
                                                        <ItemStyle HorizontalAlign="Center" />
                                                        <ItemTemplate>
                                                            <asp:Button ID="BTN_REPORT10" runat="server" Text="報表下載" CommandName="REPORT10" CssClass="asp_Export_M" />
                                                            <asp:Button ID="BTN_DOWNLOAD10" runat="server" Text="檔案下載" CommandName="DOWNLOAD10" CssClass="asp_Export_M" />
                                                            <asp:Button ID="BTN_DELFILE10" runat="server" Text="刪除檔案" CommandName="DELFILE10" CssClass="asp_button_M" />
                                                        </ItemTemplate>
                                                    </asp:TemplateColumn>
                                                </Columns>
                                                <PagerStyle Visible="False"></PagerStyle>
                                            </asp:DataGrid>
                                        </td>
                                    </tr>
                                    <tr id="tr_SENTBATVER" runat="server">
                                        <td class="bluecol">以目前版本批次送出</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Button ID="BTN_SENTBATVER" runat="server" Text="以目前版本批次送出" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr id="tr_SENDCURRVER" runat="server">
                                        <td class="bluecol">以目前版本送出</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Button ID="BTN_SENDCURRVER" runat="server" Text="以目前版本送出" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                            <%--<asp:Button ID="BTN_SENDCURRVER_DOWNLOAD" runat="server" Text="下載" CausesValidation="False" CssClass="asp_button_M"></asp:Button>--%>
                                        </td>
                                    </tr>
                                    <tr id="tr_UPLOADFL1" runat="server">
                                        <td class="bluecol">檔案上傳 </td>
                                        <td colspan="3" class="whitecol">
                                            <%--<asp:DropDownList ID="depID" runat="server"></asp:DropDownList>--%>
                                            <input id="File1" type="file" size="66" name="File1" runat="server" accept=".pdf" />
                                            <asp:Button ID="But1" runat="server" Text="確定檔案上傳" CausesValidation="False"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr id="tr_USELATESTVER" runat="server">
                                        <td class="bluecol">最近一次版本送件</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:Button ID="bt_latestSend1" runat="server" Text="最近一次版本送件" CssClass="asp_Export_M" />
                                            <asp:Button ID="bt_latestDown1" runat="server" Text="下載" CssClass="asp_Export_M" />
                                        </td>
                                    </tr>
                                    <%--WAIVED 免附文件--%>
                                    <tr id="tr_WAIVED" runat="server">
                                        <td class="bluecol">免附文件</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:CheckBox ID="CHKB_WAIVED" runat="server" />免附文件
                                        </td>
                                    </tr>
                                    <tr id="tr_USEMEMO1" runat="server">
                                        <td class="bluecol">備註說明</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="txtMEMO1" runat="server" Width="60%"></asp:TextBox>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol" colspan="4">
                                <asp:Button ID="BTN_PREV1" runat="server" Text="回上一步" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="BTN_SAVETMP1" runat="server" Text="儲存(暫存)" CssClass="asp_button_M"></asp:Button>
                                <%--<asp:Button ID="BTN_SAVERC2" runat="server" Text="下載檢核表" CssClass="asp_button_M"></asp:Button>--%>
                                <asp:Button ID="BTN_SAVENEXT1" runat="server" Text="儲存後進下一步" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="BTN_BACK1" runat="server" Text="不儲存返回查詢" CssClass="asp_button_M"></asp:Button>
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
        <input id="hid_TMID" type="hidden" name="hidReqno" runat="server" />

        <input id="hid_TPlanID28AppPlan" type="hidden" name="hid_TPlanID28AppPlan" runat="server" />
        <input id="hidReqPlanID" type="hidden" name="hidReqPlanID" runat="server" />
        <input id="hidReqcid" type="hidden" name="hidReqcid" runat="server" />
        <input id="hidReqno" type="hidden" name="hidReqno" runat="server" />

        <asp:HiddenField ID="ROC_Years" runat="server" />
        <asp:HiddenField ID="Hid_PlanYear" runat="server" />
        <asp:HiddenField ID="Hid_PCS_PR" runat="server" />
        <asp:HiddenField ID="Hid_rCDATE" runat="server" />
        <asp:HiddenField ID="Hid_SubSeqNO" runat="server" />
        <asp:HiddenField ID="Hid_ORGKINDGW" runat="server" />
        <asp:HiddenField ID="Hid_ALTDATAID" runat="server" />
        <asp:HiddenField ID="Hid_REVISESTATUS" runat="server" />
        <asp:HiddenField ID="Hid_ONLINESENDSTATUS" runat="server" />

        <asp:HiddenField ID="Hid_LastRVID" runat="server" />
        <asp:HiddenField ID="Hid_FirstRVSID" runat="server" />
        <asp:HiddenField ID="Hid_RVSID" runat="server" />
        <asp:HiddenField ID="Hid_RVID" runat="server" />
        <asp:HiddenField ID="Hid_BVFID" runat="server" />
        <asp:HiddenField ID="Hid_TECHID" runat="server" />

    </form>
</body>
</html>
