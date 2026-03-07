<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_005.aspx.vb" Inherits="WDAIIP.CO_01_005" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>審查計分表(複審)</title>
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
            // 初始化 DropDownList 的值 ddlSECONDCHK_ALL
            var ddlSECONDCHK_ALL = $(".csdatagrid1_ddlsecondchk_all");
            ddlSECONDCHK_ALL.val("");
            // 綁定事件
            ddlSECONDCHK_ALL.change(function () {
                // 獲取所有 ddlFIRSTCHK 元素
                var ddlSECONDCHKs = $("#DataGrid1").find(".cssecondchk:not(:disabled)");
                if (ddlSECONDCHKs.length == 0) { return; }
                // 遍歷所有元素
                for (var i = 0; i < ddlSECONDCHKs.length; i++) {
                    ddlSECONDCHKs[i].value = $(this).val();
                }
            });
        });

        function chackAll() {
            var Mytable = document.getElementById('DataGrid1');
            var jChoose1 = $('#Choose1');
            for (var i = 1; i < Mytable.rows.length; i++) {
                var mycheck = Mytable.rows[i].cells[0].children[0];
                if (!mycheck.disabled) {
                    mycheck.checked = jChoose1.prop("checked");//document.form1.Choose1.checked;
                }
            }
        }

        function extractDigits(inputString) {
            const str = String(inputString || ''); // Ensure it's a string, handle null/undefined
            let result = ''; // To store the extracted digits
            // Convert the string to an array of characters and then use forEach
            [...str].forEach(char => {
                // Check if the character is a digit between '0' and '9' // Append the digit to our result string
                result += (char >= '0' && char <= '9') ? char : ((char == '.') ? char : '');
            });
            return result;
        }
        function round2ODecP(num) {
            // 將數字乘以 10，然後四捨五入到最接近的整數 // 最後再除以 10，將小數點移回原位
            return Math.round(num * 10) / 10;
        }

        //部長分"changeMINISTERSUB(", iSUBTOTAL, ",$(this).val(),'", LabMINISTERSUB.ClientID, "');"
        function changeSUB1(subVal1, subVal2, labNM, txtNM, txtNM2, labNM4, oVAL) {
            // Convert inputs to numbers, defaulting to 0 if they are not valid numbers
            document.getElementById(labNM).textContent = subVal1;
            subVal2 = extractDigits(subVal2); $('#' + txtNM).val(subVal2);
            let subVal3 = $('#' + txtNM2).val();
            let numVal1 = parseFloat(subVal1) || 0;
            let numVal2 = parseFloat(subVal2) || 0;
            let numVal3 = parseFloat(subVal3) || 0;
            let onumVal = parseFloat(oVAL) || 0;
            let msg1 = "";
            $('#' + txtNM).val(numVal2);
            let max_val1 = 3;
            if (numVal2 > max_val1) {
                msg1 += ('您所輸入的分數大於最大值至多' + max_val1 + '分！`' + numVal2);
                numVal2 = max_val1; $('#' + txtNM).val(numVal2);
            } else if (numVal2 < 0) {
                msg1 += ('您所輸入的分數小於最小值至少0分！`' + numVal2);
                numVal2 = 0; $('#' + txtNM).val(numVal2);
            } else if (isNaN(numVal2) || $('#' + txtNM).val() == '') {
                msg1 += ('輸入的格式有誤應為數字格式！');
                numVal2 = 0; //$('#' + txtNM).val(numVal2); //return;
                if (onumVal > 0) { $('#' + txtNM).val(onumVal); }
            }
            if (subVal3 != '' || numVal3 > 0) {
                let numVal23 = (numVal2 + numVal3);
                if (numVal23 > 7) {
                    if (msg1 != "") { msg1 += '\n'; }
                    msg1 += ('加總的分數大於最大值至多7分！`' + numVal23);
                    numVal2 = 0; $('#' + txtNM).val('0');
                    if (onumVal > 0) { $('#' + txtNM).val(onumVal); }
                }
            }
            let val3 = numVal1 + numVal2;
            val3 = round2ODecP(val3);
            if (!isNaN(val3)) { document.getElementById(labNM).textContent = val3; }
            let val4 = numVal1 + numVal2 + numVal3;
            val4 = round2ODecP(val4);
            if (!isNaN(val4)) { document.getElementById(labNM4).textContent = val4; }
            if (msg1 != "") { alert(msg1); }
        }

        //署加分
        function changeSUB2(subVal1, subVal2, labNM, txtNM, txtNM2, oVAL) {
            // Convert inputs to numbers, defaulting to 0 if they are not valid numbers
            document.getElementById(labNM).textContent = subVal1;
            subVal2 = extractDigits(subVal2); $('#' + txtNM).val(subVal2);
            let subVal3 = $('#' + txtNM2).val();
            let numVal1 = parseFloat(subVal1) || 0;
            let numVal2 = parseFloat(subVal2) || 0;
            let numVal3 = parseFloat(subVal3) || 0;
            let onumVal = parseFloat(oVAL) || 0;
            let msg1 = "";
            $('#' + txtNM).val(numVal2);
            let max_val2 = 4;
            if (numVal2 > max_val2) {
                msg1 += ('您所輸入的分數大於最大值至多' + max_val2 + '分！`' + numVal2);
                numVal2 = max_val2; $('#' + txtNM).val(numVal2);
            } else if (numVal2 < 0) {
                msg1 += ('您所輸入的分數小於最小值至少0分！`' + numVal2);
                numVal2 = 0; $('#' + txtNM).val(numVal2);
            } else if (isNaN(numVal2) || $('#' + txtNM).val() == '') {
                msg1 += ('輸入的格式有誤應為數字格式！');
                numVal2 = 0; //$('#' + txtNM).val(numVal2); //return;
                if (onumVal > 0) { $('#' + txtNM).val(onumVal); }
            }
            if (subVal3 != '' || numVal3 > 0) {
                let numVal23 = (numVal2 + numVal3);
                if (numVal23 > 7) {
                    if (msg1 != "") { msg1 += '\n'; }
                    msg1 += ('加總的分數大於最大值至多7分！`' + numVal23);
                    numVal2 = 0; $('#' + txtNM).val('0');
                    if (onumVal > 0) { $('#' + txtNM).val(onumVal); }
                }
            }
            let val3 = numVal1 + numVal2 + numVal3;
            val3 = round2ODecP(val3);
            // Ensure val3 is a number before attempting to set the value //$('#' + labNM).val(val3); //$('#' + labNM).val('');
            if (!isNaN(val3)) { document.getElementById(labNM).textContent = val3; }
            if (msg1 != "") { alert(msg1); }
        }

        //SUBTOTAL BRANCHPNT MINISTERADD DEPTADD MINISTERSUB SCORE4_1_2,,2-部長分 3-署加分
        function changeMINISUB(typeVal) {
            //debugger; //let subVal2 = extractDigits($(this).val());
            $('#SUBTOTAL').val(extractDigits($('#SUBTOTAL').val()));
            $('#MINISTERADD').val(extractDigits($('#MINISTERADD').val()));
            $('#DEPTADD').val(extractDigits($('#DEPTADD').val()));
            $('#BRANCHPNT').val(extractDigits($('#BRANCHPNT').val()));
            let numVal1 = parseFloat($('#SUBTOTAL').val()) || 0;
            let numVal2 = parseFloat($('#MINISTERADD').val()) || 0;
            let numVal3 = parseFloat($('#DEPTADD').val()) || 0;
            let numVal4 = parseFloat($('#BRANCHPNT').val()) || 0;
            numVal1 = round2ODecP(numVal1);
            numVal2 = round2ODecP(numVal2);
            numVal3 = round2ODecP(numVal3);
            numVal4 = round2ODecP(numVal4);
            $('#SUBTOTAL').val(numVal1);
            $('#MINISTERADD').val(numVal2);
            $('#DEPTADD').val(numVal3);
            $('#BRANCHPNT').val(numVal4);

            if (typeVal == 2) {
                let max_val1 = 3;
                if (numVal2 > max_val1) {
                    //msg1 += ('您所輸入的分數大於最大值至多7分！`' + numVal2);
                    numVal2 = max_val1; $('#MINISTERADD').val(numVal2);
                } else if (numVal2 < 0) {
                    //msg1 += ('您所輸入的分數小於最小值至少0分！`' + numVal2);
                    numVal2 = 0; $('#MINISTERADD').val(numVal2);
                } else if (isNaN(numVal2) || $('#MINISTERADD').val() == '') {
                    //msg1 += ('輸入的格式有誤應為數字格式！');
                    numVal2 = 0; $('#MINISTERADD').val(numVal2); //$('#' + txtNM).val(numVal2); //return;
                }
                if ((numVal2 + numVal3) > 7) {
                    //msg1 += ('您所輸入的分數大於最大值至多7分！`' + numVal2);
                    numVal2 = round2ODecP(7 - numVal3); $('#MINISTERADD').val(numVal2);
                }
            }
            else if (typeVal == 3) {
                let max_val2 = 4;
                if (numVal3 > max_val2) {
                    //msg1 += ('您所輸入的分數大於最大值至多7分！`' + numVal2);
                    numVal3 = max_val2; $('#DEPTADD').val(numVal3);
                } else if (numVal3 < 0) {
                    //msg1 += ('您所輸入的分數小於最小值至少0分！`' + numVal2);
                    numVal3 = 0; $('#DEPTADD').val(numVal3);
                } else if (isNaN(numVal3) || $('#DEPTADD').val() == '') {
                    //msg1 += ('輸入的格式有誤應為數字格式！');
                    numVal3 = 0; $('#DEPTADD').val(numVal3); //$('#' + txtNM).val(numVal2); //return; 
                }
                if ((numVal2 + numVal3) > 7) {
                    //msg1 += ('您所輸入的分數大於最大值至多7分！`' + numVal2);
                    numVal3 = round2ODecP(7 - numVal2); $('#DEPTADD').val(numVal3);
                }
            }
            let val2 = numVal1 + numVal2;
            val2 = round2ODecP(val2);
            if (!isNaN(val2)) { $('#MINISTERSUB').val(val2); } else { $('#MINISTERSUB').val(numVal1); }
            let val3 = numVal1 + numVal2 + numVal3;
            val3 = round2ODecP(val3);
            if (!isNaN(val3)) { $('#SCORE4_1_2').val(val3); } else { $('#SCORE4_1_2').val(numVal1); }
            let val4 = numVal2 + numVal3;
            val4 = round2ODecP(val4);
            if (!isNaN(val2)) { $('#BRANCHPNT').val(val4); } else { $('#BRANCHPNT').val(''); }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;審查計分表(複審)</asp:Label>
                </td>
            </tr>
        </table>
        <%--style="display: none"--%>
        <div id="divSch1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol_need">分署
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlDISTID" runat="server"></asp:DropDownList>
                        <%--<asp:Label ID="lab_IMPDIST_MSG" runat="server" Text="(匯入必選)"></asp:Label>--%>
                    </td>
                </tr>
                <%--審查計分區間--%>
                <tr>
                    <td class="bluecol_need">審查計分區間
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlSCORING" runat="server"></asp:DropDownList>
                        <%--<asp:Label ID="lab_IMPSCORING_MSG" runat="server" Text="(匯入必選)"></asp:Label>--%>
                    </td>
                </tr>
                <%--<tr>
                    <td class="bluecol_need" style="width:20%">年度
                    </td>
                    <td class="whitecol" style="width:30%">
                        <asp:DropDownList ID="SYEARlist" runat="server"></asp:DropDownList>
                    </td>
                    <td class="bluecol_need" style="width:20%">上／下半年度</td>
                    <td class="whitecol" style="width:30%">
                        <asp:DropDownList ID="halfYear" runat="server">
                            <asp:ListItem Value="" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="1">上年度</asp:ListItem>
                            <asp:ListItem Value="2">下年度</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                </tr>--%>
                <tr>
                    <td class="bluecol" style="width: 16%">訓練機構
                    </td>
                    <td class="whitecol" style="width: 48%">
                        <asp:TextBox ID="OrgName" runat="server" MaxLength="50" Columns="60" Width="80%"></asp:TextBox>
                        <%--
                        <asp:TextBox ID="center" runat="server" Width="410px" onfocus="this.blur()"></asp:TextBox>
                        <input id="Org" type="button" value="..." name="Org" runat="server">
                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                            <input id="Orgidvalue" type="hidden" name="Orgidvalue" runat="server">
                            <span id="HistoryList2" style="position: absolute; display: none">
                            <asp:Table ID="HistoryRID" runat="server" Width="310px">
                            </asp:Table>
                        </span>--%>
                    </td>
                    <td class="bluecol" style="width: 16%">統一編號
                    </td>
                    <td class="whitecol" style="width: 20%">
                        <asp:TextBox ID="COMIDNO" runat="server" MaxLength="15" Width="50%"></asp:TextBox>
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
                <%--部-加分前：等級顯示【初擬等級】,部-加分後：等級顯示【部加分等級】,署-加分後：等級顯示【複審等級】,--%>
                <tr id="tr_rblSCORESTAGE" runat="server">
                    <td class="bluecol">(署用匯出)分數階段</td>
                    <td colspan="3" class="whitecol">
                        <asp:RadioButtonList ID="rblSCORESTAGE" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="01" Selected="True">部-加分前</asp:ListItem>
                            <asp:ListItem Value="02">部-加分後</asp:ListItem>
                            <asp:ListItem Value="03">署-加分後</asp:ListItem>
                        </asp:RadioButtonList>
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
                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button><br />
                        <%--<asp:Button ID="btnImp1" runat="server" Text="匯入總場次" CssClass="asp_button_S"></asp:Button>--%>
                        <asp:Button ID="BtnExp1" runat="server" Text="匯出審查計分表" CssClass="asp_Export_M"></asp:Button>
                        <asp:Button ID="BtnExp3" runat="server" Text="匯出統計表(署用)" CssClass="asp_Export_M"></asp:Button>
                        <asp:Button ID="BtnExp2" runat="server" Text="匯出等級比率統計表(署用)" CssClass="asp_Export_M"></asp:Button>
                        <asp:Button ID="BtnExp4" runat="server" Text="(自主)各等級分配比率" CssClass="asp_Export_M"></asp:Button>
                        <asp:Button ID="Button5" runat="server" Text="(自主)初擬等級及分數" CssClass="asp_Export_M"></asp:Button>
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
                                            <asp:TemplateColumn HeaderText="選取" HeaderStyle-Width="5%">
                                                <HeaderTemplate>
                                                    選取<input onclick="chackAll();" type="checkbox" name="Choose1" id="Choose1" />
                                                </HeaderTemplate>
                                                <ItemStyle HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <input id="checkbox1" type="checkbox" runat="server" />
                                                    <asp:HiddenField ID="HidSUBTOTAL" runat="server" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <%--<asp:BoundColumn HeaderText="序號"></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="DistName" HeaderText="分署" HeaderStyle-Width="15%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位" HeaderStyle-Width="15%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="COMIDNO" HeaderText="統一編號" HeaderStyle-Width="6%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="ORGKIND_N" HeaderText="機構別" HeaderStyle-Width="13%"></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="SUBTOTAL" HeaderText="分署<br>小計" HeaderStyle-Width="5%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <%-- <asp:BoundColumn DataField="IMPSCORE_1" HeaderText="匯入<br>分數" HeaderStyle-Width="5%"><ItemStyle HorizontalAlign="Center" /></asp:BoundColumn>--%>
                                            <%--初審等級<asp:BoundColumn DataField="LEVEL_1" HeaderText="初審<br>等級" HeaderStyle-Width="5%"><ItemStyle HorizontalAlign="Center" /></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="IMPLEVEL_1" HeaderText="初擬<br>等級" HeaderStyle-Width="5%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="ORGKIND_N" HeaderText="分署<br>加分項目"></asp:BoundColumn>--%>
                                            <asp:TemplateColumn HeaderText="部加分" HeaderStyle-Width="17%">
                                                <HeaderTemplate>
                                                    部加分<br />
                                                    <div style="display: flex; font-size: 1em; border: 0px solid black; padding: 10px;">
                                                        <span style="flex: 1; text-align: center; border-right: 1px solid white; padding-right: 5px;">加分</span>
                                                        <span style="flex: 1; text-align: center; border-right: 1px solid white; padding-right: 5px;">小計</span>
                                                        <span style="flex: 1; text-align: center;">等級</span>
                                                    </div>

                                                </HeaderTemplate>
                                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                <ItemTemplate>
                                                    <div style="display: flex; font-size: 1em;">
                                                        <span style="flex: 1; text-align: center; border-right: 1px solid white; padding-right: 5px; text-wrap: none;">
                                                            <asp:TextBox ID="tMINISTERADD" runat="server" Width="66px" MaxLength="5"></asp:TextBox></span>
                                                        <span style="flex: 1; text-align: center; border-right: 1px solid white; padding-right: 5px; text-wrap: none;">
                                                            <asp:Label ID="LabMINISTERSUB" runat="server" Text=""></asp:Label></span>
                                                        <span style="flex: 1; text-align: center; border: medium; text-wrap: none;">
                                                            <asp:DropDownList ID="ddlMINISTERLEVEL" runat="server"></asp:DropDownList></span>
                                                    </div>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="署加分" HeaderStyle-Width="7%">
                                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                                <ItemTemplate>
                                                    <asp:TextBox ID="tDEPTADD" runat="server" Width="66px" MaxLength="5"></asp:TextBox>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <%--<asp:TemplateColumn HeaderText="部加分" HeaderStyle-Width="3%">,<ItemStyle HorizontalAlign="Center" CssClass="whitecol" />,
                                               <ItemTemplate>,<asp:TextBox ID="tMINISTERADD" runat="server" Width="37px" MaxLength="5"></asp:TextBox>,</ItemTemplate>,</asp:TemplateColumn>,
                                               <asp:TemplateColumn HeaderText="部小計" HeaderStyle-Width="1%">,<ItemStyle HorizontalAlign="Center" CssClass="whitecol" />,<ItemTemplate>,
                                               <asp:Label ID="Label2" runat="server" Text=""></asp:Label>,</ItemTemplate>,</asp:TemplateColumn>,<asp:TemplateColumn HeaderText="部等級" HeaderStyle-Width="3%">,
                                               <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />,<ItemTemplate>,<asp:DropDownList ID="DropDownList1" runat="server"></asp:DropDownList>,</ItemTemplate>,
                                               </asp:TemplateColumn>--%>
                                            <%--<asp:TemplateColumn HeaderText="署／部<br>加分項目" HeaderStyle-Width="7%"><ItemStyle HorizontalAlign="Center" CssClass="whitecol" /><ItemTemplate>
                                            <asp:HiddenField ID="Hid_BRANCHPNTorg" runat="server" /><asp:TextBox ID="tBRANCHPNT" runat="server" Width="77px" MaxLength="7">
                                            </asp:TextBox></ItemTemplate></asp:TemplateColumn>--%>
                                            <%-- <asp:TemplateColumn HeaderText="小計" HeaderStyle-Width="6%"><ItemTemplate><asp:Label ID="LSUBTOTAL" runat="server"></asp:Label></ItemTemplate></asp:TemplateColumn>--%>
                                            <%--<asp:BoundColumn DataField="TOTALSCORE" HeaderText="總分" HeaderStyle-Width="5%"><ItemStyle HorizontalAlign="Center" /></asp:BoundColumn>--%>
                                            <asp:TemplateColumn HeaderText="總分" HeaderStyle-Width="5%">
                                                <ItemStyle HorizontalAlign="Center" />
                                                <ItemTemplate>
                                                    <asp:Label ID="LabTOTALSCORE" runat="server" Text=""></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="複審等級" HeaderStyle-Width="7%">
                                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                                <ItemTemplate>
                                                    <%--複審等級--<asp:Label ID="lRLEVEL_2" runat="server"></asp:Label>--%>
                                                    <asp:HiddenField ID="Hid_RLEVEL_2" runat="server" />
                                                    <asp:DropDownList ID="ddlRLEVEL_2" runat="server"></asp:DropDownList>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="複審審核" HeaderStyle-Width="7%">
                                                <HeaderTemplate>
                                                    複審審核<br />
                                                    <span class="whitecol">
                                                        <asp:DropDownList ID="ddlSECONDCHK_ALL" runat="server" CssClass="csdatagrid1_ddlsecondchk_all">
                                                            <asp:ListItem Selected="True" Value="">請選擇</asp:ListItem>
                                                            <asp:ListItem Value="Y">通過</asp:ListItem>
                                                            <asp:ListItem Value="N">不通過</asp:ListItem>
                                                        </asp:DropDownList>
                                                    </span>
                                                </HeaderTemplate>
                                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                                <ItemTemplate>
                                                    <asp:HiddenField ID="HidOSID2" runat="server" />
                                                    <asp:HiddenField ID="Hid_SECONDCHKorg" runat="server" />
                                                    <asp:DropDownList ID="ddlSECONDCHK" runat="server" CssClass="cssecondchk">
                                                        <asp:ListItem Selected="True" Value="">請選擇</asp:ListItem>
                                                        <asp:ListItem Value="Y">通過</asp:ListItem>
                                                        <asp:ListItem Value="N">不通過</asp:ListItem>
                                                    </asp:DropDownList>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="7%">
                                                <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lbtView" runat="server" Text="檢視" CommandName="btnView" CssClass="linkbutton"></asp:LinkButton>
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
                                    <asp:TextBox ID="CLSAPPCNT_t10" runat="server" Width="70px" MaxLength="7"></asp:TextBox>
                                    <%--<asp:TextBox ID="SCORE2_1_3A" runat="server" Width="70px" MaxLength="7"></asp:TextBox>--%>
                                </td>
                                <%--<td class="">=</td>
                                <td class="whitecol"><asp:TextBox ID="SCORE2_1_3_EQU" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>--%>
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
                                <%--<td class="">=</td>
                                <td class="whitecol"><asp:TextBox ID="SCORE2_2_1_EQU" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>--%>
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
                                <%--<td class="bluecol">各班分數加總</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE2_2_2_SUM" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">/</td>
                                <td class="bluecol">核定總班數</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="CLSAPPCNT_t12" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">=</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SCORE2_2_2_EQU" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>
                                <td class=""></td>--%>
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
                                    <asp:TextBox ID="SCORE3_1_N" runat="server" Width="70px" MaxLength="7"></asp:TextBox><%--牌--%></td>
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
                        <asp:TextBox ID="SCORE4_1" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
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
                                <td class="bluecol" style="width: 30%">參訓學員訓後動態調查表填寫人次</td>
                                <td class="whitecol"><%--參訓學員訓後動態調查表填寫人次--%>
                                    <asp:TextBox ID="SCORE4_2A" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">/</td>
                                <td class="bluecol">結訓學員總人次</td>
                                <td class="whitecol"><%--結訓學員總人次S--%>
                                    <asp:TextBox ID="SCORE4_2_CNT" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                                <td class="">=</td>
                                <td class="whitecol"><%--參訓學員平均填答率--%>
                                    <asp:TextBox ID="SCORE4_2_RATE" runat="server" Width="70px" MaxLength="7"></asp:TextBox>%</td>
                                <td class=""></td>
                            </tr>
                        </table>
                    </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE4_2" runat="server" Width="70px" MaxLength="7"></asp:TextBox>
                    </td>
                </tr>
                <%-- <tr><td class="bluecol" width="20%">參訓學員訓後動態調查表單位<br />
平均填答率達80% (2%)</td><td class="whitecol" width="30%">
<asp:TextBox ID="SCORE4_2_RATE" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
<td class="bluecol" width="20%">總分</td><td class="whitecol" width="30%">
<asp:TextBox ID="SCORE4_2" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td></tr>
<tr><td class="bluecol" width="25%">分署小計</td>
<td class="whitecol" colspan="3"><asp:TextBox ID="SUBTOTAL" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
<td class="bluecol" style="width: 20%">審核結果</td><td class="whitecol" style="width: 30%">
<asp:DropDownList ID="ddlFIRSTCHK_1" runat="server"><asp:ListItem Selected="True" Value="">請選擇</asp:ListItem>
<asp:ListItem Value="Y">通過</asp:ListItem><asp:ListItem Value="N">不通過</asp:ListItem></asp:DropDownList></td></tr>--%>
                <tr>
                    <td colspan="4" class="table_title_left">審查計分表(初審)</td>
                </tr>
                <%--<tr><td class="bluecol" style="width: 22%">初審審核</td><td class="whitecol" colspan="3">
<asp:DropDownList ID="ddlFIRSTCHK_1" runat="server"><asp:ListItem Selected="True" Value="">請選擇</asp:ListItem>
<asp:ListItem Value="Y">通過</asp:ListItem><asp:ListItem Value="N">不通過</asp:ListItem></asp:DropDownList></td></tr>--%>
                <%--<td class="bluecol" width="22%">初擬分數</td><td class="whitecol" width="28%">
    <asp:TextBox ID="tIMPSCORE_1" runat="server" Width="90px" MaxLength="7"></asp:TextBox></td>--%>
                <tr>
                    <td class="bluecol" style="width: 22%">分署小計</td>
                    <td class="whitecol" style="width: 28%">
                        <asp:TextBox ID="SUBTOTAL" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                    <td class="bluecol" style="width: 22%">初擬等級 </td>
                    <td class="whitecol" style="width: 28%">
                        <asp:HiddenField ID="Hid_IMPLEVEL_1" runat="server" />
                        <asp:DropDownList ID="ddlIMPLEVEL_1" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td colspan="4" class="table_title_left">5 加分項目(署)</td>
                </tr>
                <tr>
                    <%--'署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)--%>
                    <td class="bluecol">配合本部、本署辦理相關活動 或政策宣導 (7%)</td>
                    <td class="whitecol" colspan="3"><%--配合本部、本署辦理相關活動 或政策宣導 (7%)--%>
                        <asp:TextBox ID="BRANCHPNT" runat="server" Width="77px" MaxLength="5"></asp:TextBox>＝部加分<asp:TextBox ID="MINISTERADD" runat="server" Width="77px" MaxLength="7"></asp:TextBox>＋署加分<asp:TextBox ID="DEPTADD" runat="server" Width="77px" MaxLength="7"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">部加分小計</td>
                    <td class="whitecol">
                        <asp:TextBox ID="MINISTERSUB" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                    <td class="bluecol">部加分等級</td>
                    <td class="whitecol">
                        <asp:HiddenField ID="Hid_MINISTERLEVEL" runat="server" />
                        <asp:DropDownList ID="ddl_MINISTERLEVEL" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">總分</td>
                    <td class="whitecol">
                        <asp:TextBox ID="SCORE4_1_2" runat="server" Width="70px" MaxLength="7"></asp:TextBox></td>
                    <td class="bluecol">複審等級</td>
                    <td class="whitecol"><%--審查計分等級--%>
                        <asp:DropDownList ID="ddlRLEVEL_2R" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">審核結果</td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="ddlSECONDCHK_1" runat="server">
                            <asp:ListItem Selected="True" Value="">請選擇</asp:ListItem>
                            <asp:ListItem Value="Y">通過</asp:ListItem>
                            <asp:ListItem Value="N">不通過</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                </tr>
            </table>
            <table width="100%">
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

        <asp:HiddenField ID="Hid_TOTALSCORE" runat="server" />
        <asp:HiddenField ID="Hid_OSID2" runat="server" />
        <%--<asp:HiddenField ID="Hid_RTSID" runat="server" />--%>
    </form>
</body>
</html>
