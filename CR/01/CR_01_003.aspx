<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CR_01_003.aspx.vb" Inherits="WDAIIP.CR_01_003" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">

<%--<html xmlns="http://www.w3.org/1999/xhtml">--%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>審查幕僚意見</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery.blockUI.js"></script>
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
        $(document).ready(function () {
            CHK_RBL_CROSSDIST_SCH();
            // Handler for .ready() called.
            $("input[type='radio'][name='RBL_CrossDist_SCH']").on("click", function () {
                //some code
                setTimeout(function () {
                    CHK_RBL_CROSSDIST_SCH();
                }, 500);
            });
        });

        //D:不區分/C:跨區提案單位/J:轄區提案單位
        function CHK_RBL_CROSSDIST_SCH() {
            var radioValue = $("input[type='radio'][name='RBL_CrossDist_SCH']:checked").val();
            if (!radioValue) { return; }
            (radioValue == "C") ? $("#center").hide() : $("#center").show();
            (radioValue == "C") ? $("#Button2").hide() : $("#Button2").show();
            (radioValue == "C") ? $("#lab_center_msg2").show() : $("#lab_center_msg2").hide();
            //(radioValue && radioValue == "C") ? $("#HistoryList2").hide() : $("#HistoryList2").show();
            //if (radioValue) { alert("Your are a - " + radioValue); }
        }

        function SelectAllDG(fg_checked) {
            //var MyTable = document.getElementById('DataGrid1');
            //$("input:checkbox").prop('checked', Flag_checked);
            // Find all input elements with type "checkbox" and an ID starting with 'DataGrid1' and ending with 'ChkBxDG1_DEL'
            const checkboxes1 = $('input[type="checkbox"][id^="DataGrid1"][id$="ChkBxDG1_DEL"]');
            // Process the selected checkboxes
            checkboxes1.each(function () {
                if ($(this).length > 0 && !$(this)[0].disabled) {
                    $(this)[0].checked = fg_checked;
                }
                // Perform desired action on each checkbox //console.log(checkbox.attr('id')); // Example: Log the ID of each checkbox
            });
        }

        function SelectAll_ICAP(fg_checked) {
            //var MyTable = document.getElementById('DataGrid1');
            //$("input:checkbox").prop('checked', Flag_checked);
            // Find all input elements with type "checkbox" and an ID starting with 'DataGrid1' and ending with 'ChkBxDG1_DEL'
            const checkboxes1 = $('input[type="checkbox"][id^="DataGrid1"][id$="ChkBxDG1_ICAP"]');
            // Process the selected checkboxes
            checkboxes1.each(function () {
                if ($(this).length > 0 && !$(this)[0].disabled) {
                    $(this)[0].checked = fg_checked;
                }
                // Perform desired action on each checkbox //console.log(checkbox.attr('id')); // Example: Log the ID of each checkbox
            });
        }

        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;課程審查&gt;&gt;一階審查&gt;&gt;審查幕僚意見</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="PanelSch1" runat="server">
            <table width="100%" cellpadding="1" cellspacing="1" class="table_sch">
                <%--<tr><td class="bluecol_need" width="18%">年度</td><td class="whitecol" width="82%" colspan="3">
    <asp:DropDownList ID="ddlYEARS_SCH" runat="server"></asp:DropDownList></td></tr>--%>
                <tr>
                    <td class="bluecol_need" width="18%">申請階段</td>
                    <td class="whitecol" width="82%" colspan="3">
                        <asp:DropDownList ID="ddlAPPSTAGE_SCH" runat="server"></asp:DropDownList>
                        <%--<asp:Label ID="labPLANYEARS" runat="server" Text=""></asp:Label>--%></td>
                </tr>
                <%--<tr><td class="bluecol">轄區分署</td><td class="whitecol">
                   <asp:CheckBoxList ID="cblDistid" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol"></asp:CheckBoxList>
                   <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server"></td></tr>--%>
                <tr>
                    <td class="bluecol">訓練機構</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />&nbsp;
							<input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini" />
                        <span id="HistoryList2" style="position: absolute; display: none">
                            <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                        </span>
                        <asp:Label ID="lab_center_msg2" runat="server" Text="(選擇 跨區提案單位，排除【訓練機構】條件)" Style="color: #808080; display: none"></asp:Label>
                    </td>
                </tr>
                <tr id="TRPlanPoint28" runat="server">
                    <td class="bluecol">計畫</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="rblOrgKind2" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                            <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">開訓日期</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        ～<asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">&nbsp;篩選範圍</td>
                    <td colspan="3" class="whitecol">
                        <asp:RadioButtonList ID="RBL_RANGE1_SCH" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="1" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="2">轄區單位</asp:ListItem>
                            <asp:ListItem Value="3">19大類主責課程</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">&nbsp;初審建議結論</td>
                    <td colspan="3" class="whitecol">
                        <asp:RadioButtonList ID="RBL_ST1RESULT_SCH" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="1" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="2">有值</asp:ListItem>
                            <asp:ListItem Value="3">無值</asp:ListItem>
                            <asp:ListItem Value="Y">通過</asp:ListItem>
                            <asp:ListItem Value="N">不通過</asp:ListItem>
                            <asp:ListItem Value="P">調整後通過</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">跨區/轄區提案</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="RBL_CrossDist_SCH" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="D" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="C">跨區提案單位</asp:ListItem>
                            <asp:ListItem Value="J">轄區提案單位</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">iCAP課程</td>
                    <td class="whitecol" colspan="3">
                        <asp:CheckBox ID="CB1_ICAPSTAMP" runat="server" Text="僅限有iCap標章課程" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">課程申請流水號 </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="schPSNO28" runat="server" Width="30%" MaxLength="11"></asp:TextBox></td>
                </tr>
                <tr id="trBtnIMPORT1" runat="server">
                    <td class="bluecol">匯入審查幕僚意見 </td>
                    <td class="whitecol" colspan="3">
                        <%--accept=".xls,.ods"--%>
                        <asp:CheckBox ID="ChkBxCover1" runat="server" Text="確定覆蓋" />
                        <input id="File1" type="file" size="66" name="File1" runat="server" accept=".xlsx" />
                        <asp:Button ID="BtnIMPORT1" runat="server" Text="匯入" CssClass="asp_button_M"></asp:Button>(必須為xlsx格式)
                        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/CR_01_003_T1V3.zip" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" align="center" colspan="4">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="BtnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="BtnExport1" runat="server" Text="匯出提案彙整表" CssClass="asp_Export_M"></asp:Button><%--提案彙整表--%>
                        <%--<asp:Button ID="BtnExport2" runat="server" Text="匯出審查意見綜整表" CssClass="asp_Export_M"></asp:Button>--%><%--審查意見綜整表--%>
                    </td>
                </tr>
            </table>

            <div align="center">
                <asp:Label ID="msg1" runat="server" ForeColor="Red" CssClass="font"></asp:Label>
            </div>
            <table id="tbDataGrid1" runat="server" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <asp:Button ID="BtnDELChkBxDG1" runat="server" Text="勾選刪除" CssClass="asp_button_M" CommandName="BtnDelDG6" />
                        <asp:Button ID="BtnCBL_SENDICAP_RESULTS" runat="server" Text="勾選寄送iCAP課程比對結果" CssClass="asp_button_M" CommandName="BtnCBL_SENDICAP_RESULTS" />
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <%--<asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">--%>
                        <asp:DataGrid ID="DataGrid1" runat="server" AllowPaging="True" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:TemplateColumn HeaderText="刪除" HeaderStyle-VerticalAlign="Middle" HeaderStyle-HorizontalAlign="Center" HeaderStyle-Width="4%" ItemStyle-VerticalAlign="Middle" ItemStyle-HorizontalAlign="Center" ItemStyle-CssClass="whitecol">
                                    <HeaderTemplate>刪除<input onclick="SelectAllDG(this.checked);" type="checkbox" /></HeaderTemplate>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="ChkBxDG1_DEL" runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="iCAP比對" HeaderStyle-VerticalAlign="Middle" HeaderStyle-HorizontalAlign="Center" HeaderStyle-Width="4%" ItemStyle-VerticalAlign="Middle" ItemStyle-HorizontalAlign="Center" ItemStyle-CssClass="whitecol">
                                    <HeaderTemplate>iCap比對<input onclick="SelectAll_ICAP(this.checked);" type="checkbox" /></HeaderTemplate>
                                    <ItemTemplate>
                                        <asp:CheckBox ID="ChkBxDG1_ICAP" runat="server" />
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" HeaderStyle-HorizontalAlign="Center" HeaderStyle-Width="4%">
                                    <ItemTemplate>
                                        <asp:HiddenField ID="Hid_DataKey" runat="server" />
                                        <asp:HiddenField ID="Hid_DataKey2" runat="server" />
                                        <asp:Label ID="labserial_number" runat="server" Text=""></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="YEARS_ROC" HeaderText="年度" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                <asp:BoundColumn DataField="APPSTAGE_N" HeaderText="申請階段" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                <asp:BoundColumn DataField="PSNO28" HeaderText="課程申請流水號" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                <asp:BoundColumn DataField="STDATE" HeaderText="訓練起日" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                <asp:BoundColumn DataField="FTDATE" HeaderText="訓練迄日" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                <asp:BoundColumn DataField="DISTNAME" HeaderText="管控單位" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                <asp:BoundColumn DataField="ORGNAME" HeaderText="機構名稱" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                <asp:BoundColumn DataField="CLASSCNAME" HeaderText="班名" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                <asp:BoundColumn DataField="GCODEPNAME" HeaderText="職類課程" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="ST1RESULT" HeaderText="初審建議結論" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>--%>
                                <asp:TemplateColumn HeaderText="初審建議結論" ItemStyle-HorizontalAlign="Center">
                                    <ItemTemplate>
                                        <asp:Label ID="labST1RESULT_N" runat="server" Text="初審建議結論"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center">
                                    <HeaderStyle HorizontalAlign="center" Width="12%"></HeaderStyle>
                                    <ItemStyle CssClass="whitecol" />
                                    <ItemTemplate>
                                        <%--<asp:HiddenField ID="Hid_YEARS" runat="server" /><asp:HiddenField ID="Hid_APPSTAGE" runat="server" /><asp:HiddenField ID="Hid_GCODE" runat="server" /><asp:HiddenField ID="Hid_DISTID" runat="server" /><asp:RadioButtonList ID="rbl_DISTNM" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:RadioButtonList>--%>
                                        <%--<asp:TextBox ID="txtCLASSQUOTA" runat="server" MaxLength="10"></asp:TextBox>--%>
                                        <asp:Button ID="BtnADD1" runat="server" Text="新增" CommandName="ADD1" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="BtnEDT1" runat="server" Text="修改" CommandName="EDT1" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="BtnDEL1" runat="server" Text="刪除" CommandName="DEL1" CssClass="asp_button_M"></asp:Button>
                                        <asp:Button ID="BtnVIE1" runat="server" Text="查看" CommandName="VIE1" CssClass="asp_Export_M"></asp:Button>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                        <%--<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>--%>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" align="center">
                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                        <%--<br /><asp:Button ID="Button9" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>--%>
                        <%--<asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>--%>
                    </td>
                </tr>
            </table>
        </asp:Panel>

        <asp:Panel ID="PanelEdit1" runat="server">
            <table id="tbPanelEdit1" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="bluecol" width="18%">年度 </td>
                    <td class="whitecol" width="32%">
                        <asp:Label ID="lbYEARS_ROC" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol" width="18%">轄區 </td>
                    <td class="whitecol" width="32%">
                        <asp:Label ID="lbDistName" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">訓練機構 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbOrgName" runat="server"></asp:Label></td>
                    <td class="bluecol">課程申請流水號 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbPSNO28" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">班級名稱 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbClassName" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol">訓練期間 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbSFTDate" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">職類課程</td>
                    <td class="whitecol">
                        <asp:Label ID="lbGCODEPNAME" runat="server"></asp:Label>
                        <%--訓練業別<asp:Label ID="lbGCNAME" runat="server"></asp:Label>--%>
                    </td>
                    <td class="bluecol">訓練職能 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbCCNAME" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">訓練人次 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbTNum" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol">訓練時數 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbTHours" runat="server"></asp:Label>
                    </td>
                </tr>
                <%--實際人時成本、實際材料費、是否跨區提案、iCAP標章證號、政府政策性產業--%>
                <tr>
                    <td class="bluecol">實際人時成本 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbACTHUMCOST" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol">實際材料費 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbMETSUMCOST" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">是否跨區提案 </td>
                    <td class="whitecol"><%--Is it a cross-regional proposal?--%>
                        <asp:Label ID="lbIsCROSSDIST" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol">iCAP標章證號 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbiCAPNUM" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">政府政策性產業 </td>
                    <td class="whitecol" colspan="3">
                        <asp:Label ID="lbD20KNAME" runat="server"></asp:Label>
                        <asp:Label ID="lbD25KNAME" runat="server"></asp:Label>
                        <asp:Label ID="lbD26KNAME" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="table_title" colspan="4">初審幕僚建議 </td>
                </tr>
                <tr>
                    <td class="bluecol">分署幕僚意見</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="ST1SUGGEST" runat="server" TextMode="multiline" Width="80%" Rows="10"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">其他應修正內容</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="OTHFIXCONT" runat="server" TextMode="multiline" Width="80%" Rows="10"></asp:TextBox></td>
                </tr>
                <%--<tr><td class="bluecol">送請委員確認內容</td><td class="whitecol" colspan="3">
                    <asp:TextBox ID="CONFIRMCONT" runat="server" TextMode="multiline" Width="80%" Rows="10"></asp:TextBox></td></tr>--%>
                <tr>
                    <td class="bluecol">初審建議結論</td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="ddlST1RESULT" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">分署確認課程分類</td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="ddlGCODE" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" align="center" class="whitecol">
                        <asp:Button ID="btnSAVE1" runat="server" Text="確認" CssClass="asp_button_M"></asp:Button>&nbsp;
								<asp:Button ID="btnBACK1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:HiddenField ID="Hid_PSOID" runat="server" />
        <asp:HiddenField ID="Hid_PSNO28" runat="server" />
        <asp:HiddenField ID="Hid_GCODE" runat="server" />
        <asp:HiddenField ID="Hid_PFGCODE" runat="server" />
        <asp:HiddenField ID="Hid_YEARS_SCH" runat="server" />
        <asp:HiddenField ID="Hid_APPSTAGE_SCH" runat="server" />
        <asp:HiddenField ID="Hid_YN_STAFFOPINSWITCH" runat="server" />
        <asp:HiddenField ID="Hid_t_STAFFOPINSWITCH" runat="server" />
    </form>
</body>
</html>
