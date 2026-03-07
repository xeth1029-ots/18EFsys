<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_004.aspx.vb" Inherits="WDAIIP.CO_01_004" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>審查計分表(初審)</title>
    <%--<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />--%>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery.blockUI.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        $(document).ready(function () {
            //初始化 DropDownList 的值 ddlFIRSTCHK_ALL,綁定事件,初審審核
            /*var ddlFIRSTCHK_ALL = $(".csdatagrid1_ddlfirstchk_all");,ddlFIRSTCHK_ALL.val("");,ddlFIRSTCHK_ALL.change(function () {
             * ,獲取所有 ddlFIRSTCHK 元素,初審審核,var ddlFIRSTCHKs = $("#DataGrid1").find(".csfirstchk:not(:disabled)"); ddlFIRSTCHK
             * ,if (ddlFIRSTCHKs.length == 0) { return; },遍歷所有 ddlFIRSTCHK 元素,for (var i = 0; i < ddlFIRSTCHKs.length; i++) {
             * ,ddlFIRSTCHKs[i].value = $(this).val();,},});*/
            autorecsubtotal();
        });

        function chackAll() {
            var Mytable = document.getElementById('DataGrid1');
            var jChoose1 = $('#Choose1');
            for (var i = 1; i < Mytable.rows.length; i++) {
                var mycheck = Mytable.rows[i].cells[0].children[0];//選取
                if (!mycheck.disabled) {
                    mycheck.checked = jChoose1.prop("checked");//document.form1.Choose1.checked;
                }
            }
        }

        //Automatically recalculate subtotals
        function autorecsubtotal() {
            //debugger;
            $SCORE1_1 = $('#SCORE1_1');
            $SCORE1_2 = $('#SCORE1_2');
            //$SCORE2_1_1A = $('#SCORE2_1_1A');,$SCORE2_1_1B = $('#SCORE2_1_1B');,$SCORE2_1_1C = $('#SCORE2_1_1C');,$SCORE2_1_1D = $('#SCORE2_1_1D');
            $SCORE2_1_1_ALL = $('#SCORE2_1_1_ALL');

            $SCORE2_1_2_SUM_ALL = $('#SCORE2_1_2_SUM_ALL');

            $SCORE2_1_3 = $('#SCORE2_1_3');
            $SCORE2_2_1 = $('#SCORE2_2_1');
            $SCORE2_2_2 = $('#SCORE2_2_2');
            $SCORE2_3_1 = $('#SCORE2_3_1');

            $SCORE3_1 = $('#SCORE3_1');
            $SCORE3_2 = $('#SCORE3_2');
            //$SCORE4_1_A = $('#SCORE4_1_A');
            $SCORE4_1 = $('#SCORE4_1');
            $SCORE4_2 = $('#SCORE4_2');
            if ($SCORE4_1.val() != "" && $.isNumeric(parseFloat($SCORE4_1.val()))) {
                if (parseFloat($SCORE4_1.val()) > 3) { $SCORE4_1.val("3"); }
                else if (parseFloat($SCORE4_1.val()) < 0) { $SCORE4_1.val(""); }
                else if (parseFloat($SCORE4_1.val()) = 0) { $SCORE4_1.val("0"); }
            }
            else if ($SCORE4_1.val() != "" && !$.isNumeric(parseFloat($SCORE4_1.val()))) {
                $SCORE4_1.val("");
            }

            //$SCORE4_2_RATE = $('#SCORE4_2_RATE');
            //debugger;
            $('#SUBTOTAL').val("");
            var iSubTotal = parseFloat(0);
            iSubTotal += parseFloat($SCORE1_1.val());
            iSubTotal += parseFloat($SCORE1_2.val());
            //iSubTotal += parseFloat($SCORE2_1_1A.val());,iSubTotal += parseFloat($SCORE2_1_1B.val());,iSubTotal += parseFloat($SCORE2_1_1C.val());,iSubTotal += parseFloat($SCORE2_1_1D.val());
            iSubTotal += parseFloat($SCORE2_1_1_ALL.val());
            iSubTotal += parseFloat($SCORE2_1_2_SUM_ALL.val());
            //iSubTotal += parseFloat($SCORE2_1_2A.val());,iSubTotal += parseFloat($SCORE2_1_2B.val());,iSubTotal += parseFloat($SCORE2_1_2C.val());,iSubTotal += parseFloat($SCORE2_1_2D.val());
            iSubTotal += parseFloat($SCORE2_1_3.val());
            iSubTotal += parseFloat($SCORE2_2_1.val());
            iSubTotal += parseFloat($SCORE2_2_2.val());
            iSubTotal += parseFloat($SCORE2_3_1.val());

            iSubTotal += parseFloat($SCORE3_1.val());
            iSubTotal += parseFloat($SCORE3_2.val());

            if ($SCORE4_1.val() != "" && $.isNumeric(parseFloat($SCORE4_1.val()))) { iSubTotal += parseFloat($SCORE4_1.val()); }
            if ($SCORE4_2.val() != "" && $.isNumeric(parseFloat($SCORE4_2.val()))) { iSubTotal += parseFloat($SCORE4_2.val()); }
            //if ($SCORE4_2_RATE.val() != "") { iSubTotal += parseFloat($SCORE4_2_RATE.val()); }
            if ($('#SUBTOTAL').val() == "") { $('#SUBTOTAL').val(parseFloat(iSubTotal).toFixed(2)); }
            //console.log("autorecsubtotal()");
        }
        function click_btnResetSUBTOTAL() {
            //console.log("click_btnResetSUBTOTAL()");
            autorecsubtotal();
            return false;
        }
        function AUTO_CAL_2(oSCORE4_org, oSCORE4, oSUBTOTA_org, oSUBTOTA) {
            //Hid_SCORE4_1org.ClientID, tSCORE4_1.ClientID, Hid_SUBTOTALorg.ClientID, tSUBTOTAL.ClientID
            $SCORE4_org = $("#" + oSCORE4_org);
            $SCORE4 = $("#" + oSCORE4);
            $SUBTOTA_org = $("#" + oSUBTOTA_org);
            $SUBTOTA = $("#" + oSUBTOTA);
            //加總非數字，無法計算
            if (isNaN($SUBTOTA_org.val())) { return false; }
            //取得加總
            var iSubTotal = parseFloat($SUBTOTA_org.val());
            if ($SCORE4.val() != "" && isNaN($SCORE4.val())) {
                alert("加分項目由分署自填，請填寫數字!"); //非數字
                $SCORE4.val("");
                return false;
            }
            else if ($SCORE4.val() != "" && parseFloat($SCORE4.val()) > 3) {
                alert("加分項目由分署自填，請填寫數字，至多加 3 分!"); //數字太大
                $SCORE4.val("3");
                return false;
            }
            else if ($SCORE4.val() != "" && parseFloat($SCORE4.val()) < 0) {
                alert("加分項目由分署自填，請填寫數字，至少加 0 分!"); //數字太小
                $SCORE4.val("0");
                return false;
            }
            //可計算
            if ($SCORE4_org.val() != "" && !isNaN($SCORE4_org.val())) { iSubTotal = iSubTotal - parseFloat($SCORE4_org.val()); }
            if ($SCORE4.val() != "" && !isNaN($SCORE4.val())) {
                iSubTotal = iSubTotal + (parseFloat($SCORE4.val()) > 3 ? 3 : parseFloat($SCORE4.val()));
            }
            //結束
            $SUBTOTA.val(parseFloat(iSubTotal).toFixed(2));
            return false;
        }
    </script>
    <%--<style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; width: 20%; height: 29px; }
        .auto-style2 { color: #333333; padding: 4px; height: 29px; }
    </style>--%>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;審查計分表(初審)</asp:Label>
                </td>
            </tr>
        </table>
        <%--style="display: none"--%>
        <div id="divSch1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol_need" style="width: 20%">分署
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlDISTID" runat="server"></asp:DropDownList>
                        <asp:Label ID="lab_IMPDIST_MSG" runat="server" Text="(匯入必選)"></asp:Label>
                    </td>
                </tr>
                <%--審查計分區間--%>
                <tr>
                    <td class="bluecol_need" style="width: 20%">審查計分區間
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlSCORING" runat="server"></asp:DropDownList>
                        <asp:Label ID="lab_IMPSCORING_MSG" runat="server" Text="(匯入必選)"></asp:Label>
                    </td>
                </tr>
                <%--<tr><td class="bluecol_need" style="width:20%">年度</td>
<td class="whitecol" style="width:30%"><asp:DropDownList ID="SYEARlist" runat="server"></asp:DropDownList>
</td><td class="bluecol_need" style="width:20%">上／下半年度</td><td class="whitecol" style="width:30%">
<asp:DropDownList ID="halfYear" runat="server"><asp:ListItem Value="" Selected="True">不區分</asp:ListItem>
<asp:ListItem Value="1">上年度</asp:ListItem><asp:ListItem Value="2">下年度</asp:ListItem></asp:DropDownList></td></tr>--%>
                <tr>
                    <td class="bluecol">訓練機構
                    </td>
                    <td class="whitecol" style="width: 35%">
                        <asp:TextBox ID="OrgName" runat="server" MaxLength="50" Columns="60" Width="90%"></asp:TextBox>
                    </td>
                    <td class="bluecol" style="width: 20%">統一編號
                    </td>
                    <td class="whitecol" style="width: 25%">
                        <asp:TextBox ID="COMIDNO" runat="server" MaxLength="15" Width="60%"></asp:TextBox>
                    </td>
                </tr>
                <tr id="TRPlanPoint28" runat="server">
                    <td class="bluecol_need">計畫</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="OrgPlanKind" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                            <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">機構別 </td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="OrgKindList" runat="server" CssClass="font"></asp:DropDownList>
                    </td>
                </tr>
                <%-- <tr>,<td class="bluecol">初審審核狀態</td>,<td colspan="3" class="whitecol">,<asp:RadioButtonList ID="rblFIRSTCHK_SCH" runat="server" RepeatDirection="Horizontal" Width="33%">
                ,<asp:ListItem Selected="True" Value="A">不區分</asp:ListItem>,<asp:ListItem Value="Y">通過</asp:ListItem>,<asp:ListItem Value="N">不通過</asp:ListItem>,</asp:RadioButtonList>,</td>,</tr>--%>
                <%-- <tr><td class="bluecol">評核版本</td><td class="whitecol"><asp:DropDownList ID="ddlSENDVER" runat="server" CssClass="font"></asp:DropDownList></td>
                <td class="bluecol">評核結果</td><td class="whitecol"><asp:DropDownList ID="ddlRESULT" runat="server" CssClass="font"></asp:DropDownList></td></tr>--%>
                <tr id="trImport1" runat="server">
                    <td class="bluecol">匯入等級/分數 </td>
                    <td class="whitecol" colspan="3">
                        <input id="File1" type="file" size="55" name="File1" runat="server" accept=".xls,.ods" />
                        <asp:Button ID="btnImport1" runat="server" Text="匯入等級/分數" CssClass="asp_button_M"></asp:Button>(必須為ods或xls格式)
                        <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../Doc/Co_Temp_v1.zip" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                    </td>
                </tr>
                <tr id="trImport2" runat="server">
                    <td class="bluecol">匯入分署加分</td>
                    <td class="whitecol" colspan="3">
                        <input id="File2" type="file" size="55" name="File2" runat="server" accept=".xls,.ods" />
                        <asp:Button ID="btnImport2" runat="server" Text="匯入分署加分" CssClass="asp_button_M"></asp:Button>(必須為ods或xls格式)
                        <asp:HyperLink ID="HyperLink2" runat="server" NavigateUrl="../../Doc/Co_Temp_v2.zip" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
                    </td>
                </tr>
                <tr id="trImport3" runat="server">
                    <td class="bluecol">匯入初擬等級</td>
                    <td class="whitecol" colspan="3">
                        <input id="File3" type="file" size="55" name="File3" runat="server" accept=".xls,.ods" />
                        <asp:Button ID="btnImport3" runat="server" Text="匯入初擬等級" CssClass="asp_button_M"></asp:Button>(必須為ods或xls格式)
                        <asp:HyperLink ID="HyperLink3" runat="server" NavigateUrl="../../Doc/Co_Temp_v3.zip" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>
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
                    <td class="whitecol" colspan="4" align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <%--<asp:Button ID="btnImp1" runat="server" Text="匯入總場次" CssClass="asp_button_S"></asp:Button>--%>
                        <asp:Button ID="btnExp1" runat="server" Text="匯出審查計分表" CssClass="asp_Export_M" data-exp="Y"></asp:Button>
                        <asp:Button ID="btnExp2" runat="server" Text="匯出單位計分" CssClass="asp_Export_M" data-exp="Y2"></asp:Button><br />
                        <asp:Button ID="btnExp3" runat="server" Text="匯出班級明細計分" CssClass="asp_Export_M" data-exp="Y3"></asp:Button>
                        <asp:Button ID="btnExp4" runat="server" Text="匯出統計表" CssClass="asp_Export_M" data-exp="Y4"></asp:Button>
                        <asp:Button ID="btnExp5" runat="server" Text="匯出等級比率統計表" CssClass="asp_Export_M" data-exp="Y5"></asp:Button>
                    </td>
                </tr>
            </table>
            <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center">
                        <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left"
                                        PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="選取" HeaderStyle-Width="4%">
                                                <HeaderTemplate>
                                                    選取<input onclick="chackAll();" type="checkbox" name="Choose1" id="Choose1" />
                                                </HeaderTemplate>
                                                <ItemStyle HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <input id="checkbox1" type="checkbox" runat="server" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <%--<asp:BoundColumn HeaderText="序號"></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="DistName" HeaderText="分署" HeaderStyle-Width="11%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位" HeaderStyle-Width="15%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="COMIDNO" HeaderText="統一編號" HeaderStyle-Width="6%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ORGKIND_N" HeaderText="機構別" HeaderStyle-Width="10%"></asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="ORGKIND_N" HeaderText="分署<br>加分項目"></asp:BoundColumn>--%>
                                            <asp:TemplateColumn HeaderText="分署加分項目" HeaderStyle-Width="7%">
                                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                                <ItemTemplate>
                                                    <%--<asp:HiddenField ID="Hid_BRANCHPNTorg" runat="server" />
                                                    <asp:TextBox ID="tBRANCHPNT" runat="server" Width="60%" MaxLength="7" size="2"></asp:TextBox>--%>
                                                    <asp:HiddenField ID="Hid_SCORE4_1org" runat="server" />
                                                    <asp:TextBox ID="tSCORE4_1" runat="server" Width="95%" MaxLength="7"></asp:TextBox>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="小計" HeaderStyle-Width="7%">
                                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                                <ItemTemplate>
                                                    <%--<asp:Label ID="LSUBTOTAL" runat="server"></asp:Label>--%>
                                                    <asp:HiddenField ID="Hid_SUBTOTALorg" runat="server" />
                                                    <asp:TextBox ID="tSUBTOTAL" runat="server" Width="95%" MaxLength="7"></asp:TextBox>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="SORTRATIO1_N" HeaderText="排序<br>比率" HeaderStyle-Width="6%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="SORTLEVEL1" HeaderText="系統排序等級" HeaderStyle-Width="6%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="IMPLEVEL_1" HeaderText="初擬<br>等級" HeaderStyle-Width="5%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="說明" HeaderStyle-Width="6%">
                                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                                <ItemTemplate>
                                                    <asp:Label ID="labCAPIDX1" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lbtView" runat="server" Text="檢視" CommandName="btnView" CssClass="linkbutton"></asp:LinkButton>
                                                    <asp:HiddenField ID="HidOSID2" runat="server" />
                                                    <%--<asp:HiddenField ID="Hid_FIRSTCHKorg" runat="server" />--%>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <%--<asp:BoundColumn DataField="MEMO" HeaderText="備註"></asp:BoundColumn>--%>
                                        </Columns>
                                        <PagerStyle Visible="False" HorizontalAlign="Left" ForeColor="Blue" Position="Top" Mode="NumericPages"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>

                        </table>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Label ID="Labmsg2" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol">
                        <div align="center">
                            <%--<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="55px">10</asp:TextBox>
                            <asp:Button ID="BtnBack1" runat="server" Text="回上頁" CssClass="asp_button_S"></asp:Button>--%>
                            <asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>

            </table>
        </div>
        <%--style="display: none"--%>
        <div id="divEdt1" runat="server">
            <table class="table_nw" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol">訓練單位</td>
                    <td class="whitecol" colspan="3">
                        <asp:Label ID="LabOrgName" runat="server"></asp:Label>
                    </td>
                </tr>
                <%--審查計分區間--%>
                <tr>
                    <td class="bluecol" style="width: 20%">審查計分區間
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:Label ID="labSCORING_N" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 20%">分署
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:Label ID="LabDISTNAME" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr id="tr_Lab_SUSPENDED_msg1" runat="server">
                    <%--<td class="bluecol" style="width: 20%"></td>--%>
                    <td colspan="4" class="whitecol">
                        <asp:Label ID="Lab_SUSPENDED_msg1" runat="server" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="table_title_left">1-1 開班率(5%)</td>
                </tr>
                <tr>
                    <td colspan="3" class="class_title3" width="85%">審查計分公式</td>
                    <td class="class_title3" width="15%">得分</td>
                </tr>
                <tr>
                    <td colspan="3" class="">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="whitecol" colspan="2">
                                    <table cellpadding="0" cellspacing="0" width="100%" border="0">
                                        <tr>
                                            <td class="whitecol">（</td>
                                            <td class="bluecol">實際開班數</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="CLSACTCNT" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                            <td class="whitecol">＋</td>
                                            <td class="bluecol">政策性課程班數</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="CLSACTCNT2" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                            <td class="whitecol">）</td>
                                        </tr>
                                    </table>
                                </td>
                                <td class="">/</td>
                                <td class="bluecol">核定總班數</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="CLSAPPCNT" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">=</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE1_1A" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>
                                <td class=""></td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE1_1" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                </tr>
                <tr>
                    <td colspan="4" class="table_title_left">1-2 訓練人次達成率(8%)</td>
                </tr>
                <tr>
                    <td colspan="3" class="class_title3">審查計分公式</td>
                    <td class="class_title3">得分</td>
                </tr>
                <tr>
                    <td colspan="3" class="">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="whitecol" colspan="2">
                                    <table cellpadding="0" cellspacing="0" width="100%" border="0">
                                        <tr>
                                            <td class="whitecol">（</td>
                                            <td class="bluecol">實際開訓人次</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="STDACTCNT" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                            <td class="whitecol">＋</td>
                                            <td class="bluecol">政策性課程核定人次</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="STDACTCNT2" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                            <td class="whitecol">）</td>
                                        </tr>
                                    </table>
                                </td>
                                <td class="">/</td>
                                <td class="bluecol">核定總人次</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="STDAPPCNT" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">=</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE1_2A" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>
                                <td class=""></td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE1_2" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                </tr>
                <tr>
                    <td colspan="4" class="table_title_left">2-1 資料建置及維護(28%) </td>
                </tr>
                <tr>
                    <td colspan="4" class="class_title2_left">2-1-1 各項函送資料及資訊登錄作業時效(11%)</td>
                </tr>
                <tr>
                    <td colspan="4">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="class_title3">項目</td>
                                <td class="class_title3">審查計分公式</td>
                                <td class="class_title3">核定總班數</td>
                                <td class="class_title3">得分</td>
                            </tr>
                            <tr>
                                <td class="">招訓資料</td>
                                <td class="">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol">各班分數加總</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="SCORE2_1_1_SUM_A" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                                <td class="whitecol" rowspan="4" valign="middle" align="center">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="whitecol" valign="middle" align="right">/</td>
                                            <td class="whitecol" valign="middle" align="left">
                                                <asp:TextBox ID="CLSAPPCNT_t2" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                                <td class="whitecol" rowspan="4" valign="middle" align="center">
                                    <asp:TextBox ID="SCORE2_1_1_ALL" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="">開訓資料</td>
                                <td class="">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol">各班分數加總</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="SCORE2_1_1_SUM_B" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                                <%--<td class="whitecol"><asp:TextBox ID="SCORE2_1_1B" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>--%>
                            </tr>
                            <tr>
                                <td class="">結訓資料</td>
                                <td class="">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol">各班分數加總</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="SCORE2_1_1_SUM_C" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                                <%--<td class="whitecol"><asp:TextBox ID="SCORE2_1_1C" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>--%>
                            </tr>
                            <tr>
                                <td class="">變更申請</td>
                                <td class="">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol">各班分數加總</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="SCORE2_1_1_SUM_D" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                                <%--<td class="whitecol"><asp:TextBox ID="SCORE2_1_1D" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>--%>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="class_title2_left">2-1-2 函送資料內容及資訊登錄正確性(10%)</td>
                </tr>
                <tr>
                    <td colspan="4">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="class_title3">項目</td>
                                <td class="class_title3">審查計分公式</td>
                                <td class="class_title3">得分</td>
                            </tr>
                            <tr>
                                <td class="">招訓資料</td>
                                <td class="">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol">各班總扣分</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="SCORE2_1_2A_DIS" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                                <td class="whitecol" rowspan="4" valign="middle" align="center">
                                    <asp:TextBox ID="SCORE2_1_2_SUM_ALL" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <%--<td class="whitecol" valign="middle" align="center"><asp:TextBox ID="SCORE2_1_2A" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>--%>
                            </tr>
                            <tr>
                                <td class="">開訓資料</td>
                                <td class="">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol">各班總扣分</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="SCORE2_1_2B_DIS" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td class="">結訓資料</td>
                                <td class="">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol">各班總扣分</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="SCORE2_1_2C_DIS" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td class="">變更申請</td>
                                <td class="">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol">各班總扣分</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="SCORE2_1_2D_DIS" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>

                        </table>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="class_title2_left">2-1-3 訓練計畫變更項次數(7%)</td>
                </tr>
                <tr>
                    <td colspan="3" class="class_title3">審查計分公式</td>
                    <td class="class_title3">得分</td>
                </tr>
                <tr>
                    <td colspan="3" class="">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol">各班分數加總</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE2_1_3_SUM" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">/</td>
                                <td class="bluecol">核定總班數</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="CLSAPPCNT_t10" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <%--<asp:TextBox ID="SCORE2_1_3A" runat="server" Width="70px" MaxLength="7"></asp:TextBox>--%>
                                <%--<td class="">=</td><td class="whitecol"><asp:TextBox ID="SCORE2_1_3_EQU" runat="server" Width="70px" MaxLength="7"></asp:TextBox>(得分)</td>--%>
                                <td class=""></td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE2_1_3" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                </tr>
                <tr>
                    <td colspan="4" class="table_title_left">2-2 督導與考核(34%)</td>
                </tr>
                <tr>
                    <td colspan="4" class="class_title2_left">2-2-1 學員管理(4%)</td>
                </tr>
                <tr>
                    <td colspan="3" class="class_title3">審查計分公式</td>
                    <td class="class_title3">得分</td>
                </tr>
                <tr>
                    <td colspan="3" class="">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol">各班分數加總</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE2_2_1_SUM" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">/</td>
                                <td class="bluecol">核定總班數</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="CLSAPPCNT_t11" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <%--<td class="">=</td><td class="whitecol"><asp:TextBox ID="SCORE2_2_1_EQU" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>--%>
                                <td class=""></td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE2_2_1" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                </tr>
                <tr>
                    <td colspan="4" class="class_title2_left">2-2-2 課程辦理情形(30%)</td>
                </tr>
                <tr>
                    <td colspan="3" class="class_title3">審查計分公式</td>
                    <td class="class_title3">得分</td>
                </tr>
                <tr>
                    <td colspan="3" class="">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <%--<td class="bluecol">各班分數加總</td><td class="whitecol"><asp:TextBox ID="SCORE2_2_2_SUM" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td><td class="">/</td>
<td class="bluecol">核定總班數</td><td class="whitecol"><asp:TextBox ID="CLSAPPCNT_t12" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td><td class="">=</td>
<td class="whitecol"><asp:TextBox ID="SCORE2_2_2_EQU" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td><td class=""></td>--%>
                                <td class="bluecol">各班總扣分</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE2_2_2_DIS" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE2_2_2" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                </tr>
                <tr>
                    <td colspan="4" class="table_title_left">2-3 計畫參與度(5%)</td>
                </tr>
                <tr>
                    <td colspan="3" class="class_title3">審查計分公式</td>
                    <td class="class_title3">得分</td>
                </tr>
                <tr>
                    <td colspan="3" class="">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol">實際出席總場次</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE2_3_1_SUM" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">/</td>
                                <td class="bluecol">應出席總場次</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE2_3_1_CNT" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">=</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE2_3_1_EQU" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>
                                <td class=""></td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE2_3_1" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                </tr>
                <tr>
                    <td colspan="4" class="table_title_left">3-1 最近一次TTQS評核結果等級(10%)</td>
                </tr>
                <tr>
                    <td colspan="3" class="class_title3">審查計分公式</td>
                    <td class="class_title3">得分</td>
                </tr>
                <tr>
                    <td colspan="3" class="">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol" width="25%">訓綀機構版</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE3_1_N" runat="server" Width="70px" MaxLength="7"></asp:TextBox><%--牌--%>
                                    <asp:Label ID="labSCORE3_1_N" runat="server" Text=""></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE3_1" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                </tr>
                <tr>
                    <td colspan="4" class="table_title_left">3-2 學員滿意程度(5%)</td>
                </tr>
                <tr>
                    <td colspan="3" class="class_title3">審查計分公式</td>
                    <td class="class_title3">得分</td>
                </tr>
                <tr>
                    <td colspan="3" class="">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol">滿意學員人次</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE3_2_SUM" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">/</td>
                                <td class="bluecol">結訓學員總人次</td>
                                <td class="whitecol"><%--結訓學員總人次*2--%>
                                    <asp:TextBox ID="SCORE3_2_CNT" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">/2=</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE3_2_EQU" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>
                                <td class=""></td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE3_2" runat="server" Width="70px" MaxLength="7"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="table_title_left">4 加分項目(分署)</td>
                </tr>
                <tr>
                    <td class="bluecol" width="20%">配合分署辦理相關活動或政策宣導(3%)</td>
                    <td class="whitecol" colspan="3">
                        <%--配合分署辦理相關活動或政策宣導--%>
                        <asp:TextBox ID="SCORE4_1" runat="server" Width="70px" MaxLength="7"></asp:TextBox>
                        (請填寫數字，至多加3分，輸入後按Tab鍵或失去焦點會自動計算。)
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="class_title2_left">參訓學員訓後動態調查表單位平均填答率達80% (2%)</td>
                </tr>
                <tr>
                    <td colspan="3" class="class_title3">審查計分公式</td>
                    <td class="class_title3">得分</td>
                </tr>
                <tr>
                    <td colspan="3" class="">
                        <table cellpadding="1" cellspacing="1" width="100%" border="0">
                            <tr>
                                <td class="bluecol" width="33%">參訓學員訓後動態調查表填寫人次</td>
                                <td class="whitecol">
                                    <%--參訓學員訓後動態調查表填寫人次--%>
                                    <asp:TextBox ID="SCORE4_2A" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">/</td>
                                <td class="bluecol">結訓學員總人次</td>
                                <td class="whitecol">
                                    <%--結訓學員總人次--%>
                                    <asp:TextBox ID="SCORE4_2_CNT" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">=</td>
                                <td class="whitecol">
                                    <%--參訓學員平均填答率--%>
                                    <asp:TextBox ID="SCORE4_2_RATE" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>
                                <td class=""></td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE4_2" runat="server" Width="70px" MaxLength="7"></asp:TextBox>
                    </td>
                </tr>
                <%-- <tr><td class="bluecol" width="20%">參訓學員訓後動態調查表單位<br />平均填答率達80% (2%)</td>
                <td class="whitecol" width="30%"><asp:TextBox ID="SCORE4_2_RATE" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                <td class="bluecol" width="20%">總分</td><td class="whitecol" width="30%">
                <asp:TextBox ID="SCORE4_2" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td></tr>--%>
                <tr>
                    <td class="bluecol" width="22%">小計</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="SUBTOTAL" runat="server" Width="90px" MaxLength="7"></asp:TextBox>
                        <input id="btnResetSUBTOTAL" runat="server" type="button" value="重新計算" />
                    </td>
                </tr>
                <%-- <tr><td class="bluecol" style="width: 22%">初審審核</td><td class="whitecol" colspan="3"><asp:DropDownList ID="ddlFIRSTCHK_1" runat="server">
                   <asp:ListItem Selected="True" Value="">請選擇</asp:ListItem><asp:ListItem Value="Y">通過</asp:ListItem><asp:ListItem Value="N">不通過</asp:ListItem></asp:DropDownList></td></tr>--%>
                <tr>
                    <td class="bluecol" style="width: 22%">初擬等級 </td>
                    <td class="whitecol" colspan="3">
                        <asp:HiddenField ID="Hid_IMPLEVEL_1" runat="server" />
                        <asp:DropDownList ID="ddlIMPLEVEL_1" runat="server">
                        </asp:DropDownList>
                    </td>
                </tr>
                <%--<tr><td class="bluecol" width="25%">匯入分數</td><td class="whitecol" width="35%"><asp:TextBox ID="tIMPSCORE_1" runat="server" Width="90px" MaxLength="7">
                </asp:TextBox></td><td class="whitecol" colspan="2"></td></tr>--%>
                <%--<tr><td colspan="4" class="table_title_left">5 加分項目(署)</td></tr><tr><td class="bluecol" width="20%">配合本部、本署辦理相關活動 或政策宣導 (7%)</td><td class="whitecol">
                <asp:TextBox ID="SCORE5_1A" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td><td class="bluecol" width="20%">總分</td><td class="whitecol" width="30%">
                <asp:TextBox ID="SCORE4_1_2" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td></tr>--%>
                <%--<tr><td class="" colspan="2"></td></tr>--%>
            </table>
            <table width="100%">
                <tr>
                    <td align="center">
                        <asp:Label ID="labmsg3" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol">
                        <div align="center">
                            <%--<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="55px">10</asp:TextBox>--%>
                            <asp:Button ID="BtnBack2" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="BtnSaveData2" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
        <asp:HiddenField ID="Hid_REMIND1" runat="server" />
        <asp:HiddenField ID="Hid_OSID2" runat="server" />
        <asp:HiddenField ID="Hid_RLEVEL_2" runat="server" />
        <asp:HiddenField ID="Hid_SECONDCHK" runat="server" />
        <asp:HiddenField ID="Hid_MINISTERLEVEL" runat="server" />
        <%--<asp:HiddenField ID="Hid_RTSID" runat="server" />--%>
    </form>
</body>
</html>
