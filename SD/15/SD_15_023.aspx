<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_023.aspx.vb" Inherits="WDAIIP.SD_15_023" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>綜合動態報表-交叉分析統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function OpenOrg(vTPlanID) {
            var DistID = document.getElementById('DistID');
            if (DistID.selectedIndex == 0) {
                alert('請先選擇轄區');
                return false;
            }
            else {
                wopen('../../common/MainOrg.aspx?DistID=' + DistID.value + '&TPlanID=' + vTPlanID, '', 600, 500, 'yes');
            }
        }

        function CheckSearch() {
            var STDate1 = document.getElementById('STDate1').value;
            var STDate2 = document.getElementById('STDate2').value;

            var msg = '';
            if (!checkDate(STDate1) && STDate1 != '') msg += '開訓起始日期必須為正確日期格式\n';
            if (!checkDate(STDate2) && STDate2 != '') msg += '開訓結束日期必須為正確日期格式\n';

            if (msg != '') {
                alert(msg);
                return false;
            }

            if (!isChecked(document.form1.StudStatus)) {
                alert('請選擇統計範圍');
                return false;
            }

            if (!isChecked(document.form1.XRoll) || !isChecked(document.form1.YRoll)) {
                alert('請選擇XY軸分析項目');
                return false;
            }
        }

        //$(document).ready(function () {
        //    $("#StudStatus").change(function () {
        //        chk_StudStatus();
        //    });
        //});

        function chk_StudStatus() {
            //debugger;
            //var v_s1 = $("#StudStatus").val();
            //var v_s1 = $("select[name='StudStatus']").val();
            if (!window.jQuery) { return; }
            var v_s1 = $('input[name=StudStatus]:checked').val();
            //var tdBudgetList2 = document.getElementById("tdBudgetList2");
            var spBudgetList1 = document.getElementById("spBudgetList1"); //span
            var spBudgetList2 = document.getElementById("spBudgetList2"); //span
            if (v_s1 == undefined) { return; }
            if (tdBudgetList2 == undefined) { return; }
            if (!tdBudgetList2) { return; }
            //$("input.GBudgetList").removeAttr("disabled");
            //$("#trBudgetList").removeAttr("disabled");
            //tdBudgetList2.disabled = false;
            //tdBudgetList2.style.display = '';
            spBudgetList1.style.display = '';
            spBudgetList2.style.display = 'none';
            $("input[type='checkbox'][name^='BudgetList']").removeAttr("disabled");
            if (v_s1 == "11" || v_s1 == "") {
                //$("input.GBudgetList").attr("disabled", true);
                $("input[type='checkbox'][name^='BudgetList']").prop("checked", false);
                $("input[type='checkbox'][name^='BudgetList']").attr("disabled", true);
                //$("#trBudgetList").attr("disabled", true);
                //tdBudgetList2.disabled = true;
                //tdBudgetList2.style.display = 'none';
                spBudgetList1.style.display = 'none';
                spBudgetList2.style.display = '';
            }
        }

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
                    var mycheck = document.getElementById(obj + '_' + i);
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }


    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;綜合動態報表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="myTable1" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <table id="Table2" class="table_sch">
                        <tr>
                            <td class="bluecol" width="20%">動態報表 </td>
                            <td class="whitecol" width="80%">
                                <uc1:WUC1 runat="server" ID="WUC1" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">轄區 </td>
                            <td class="whitecol" width="80%">
                                <asp:DropDownList ID="DistID" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="Button3" value="..." type="button" name="Button1" runat="server" class="button_b_Mini" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="PlanID" type="hidden" name="PlanID" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">開訓期間 </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="STDate1" runat="server" Columns="14" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></span>～
                                <asp:TextBox ID="STDate2" runat="server" Columns="14" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">結訓期間 </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="14" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></span>～
                                <asp:TextBox ID="FTDate2" runat="server" Columns="14" MaxLength="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr id="trPlanKind" runat="server">
                            <td class="bluecol" width="20%">計畫範圍 </td>
                            <td class="whitecol" colspan="4" width="80%">
                                <asp:RadioButtonList ID="SearchPlan" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cblAPPSTAGE" runat="server" RepeatDirection="Horizontal">
                                </asp:CheckBoxList>
                                <input id="cblAPPSTAGE_Hidden" type="hidden" value="0" runat="server" name="cblAPPSTAGE_Hidden" />
                            </td>
                        </tr>
                        <tr id="trPackageType" runat="server">
                            <td class="bluecol" width="20%">包班種類 </td>
                            <td class="whitecol" colspan="4" width="80%">
                                <asp:RadioButtonList ID="PackageType" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="A" Selected="True">全部</asp:ListItem>
                                    <asp:ListItem Value="2">企業包班</asp:ListItem>
                                    <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">統計範圍 </td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="StudStatus" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellSpacing="1" CellPadding="1">
                                    <asp:ListItem Value="11" Selected="True">報名人數</asp:ListItem>
                                    <asp:ListItem Value="12">參訓人數</asp:ListItem>
                                    <asp:ListItem Value="13">結訓人數</asp:ListItem>
                                    <asp:ListItem Value="14">撥款人數</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">X軸 </td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="XRoll" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellSpacing="1" CellPadding="1">
                                    <asp:ListItem Value="1">性別</asp:ListItem>
                                    <asp:ListItem Value="2">年齡</asp:ListItem>
                                    <asp:ListItem Value="3">教育程度</asp:ListItem>
                                    <asp:ListItem Value="4">特定對象</asp:ListItem>
                                    <asp:ListItem Value="5">結訓後動向</asp:ListItem>
                                    <asp:ListItem Value="6">工作年資</asp:ListItem>
                                    <asp:ListItem Value="7">受訓學員地理分布</asp:ListItem>
                                    <asp:ListItem Value="8">所屬公司行業別</asp:ListItem>
                                    <asp:ListItem Value="9">所屬公司規模</asp:ListItem>
                                    <asp:ListItem Value="10">參訓動機</asp:ListItem>
                                    <asp:ListItem Value="15">參訓單位類別</asp:ListItem>
                                    <asp:ListItem Value="16">職能課程分類</asp:ListItem>
                                    <asp:ListItem Value="17">參加課程型態</asp:ListItem>
                                    <asp:ListItem Value="18">外籍配偶類別</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">Y軸 </td>
                            <td class="whitecol" width="80%">
                                <asp:RadioButtonList ID="YRoll" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatColumns="4" CellSpacing="1" CellPadding="1">
                                    <asp:ListItem Value="1">性別</asp:ListItem>
                                    <asp:ListItem Value="2">年齡</asp:ListItem>
                                    <asp:ListItem Value="3">教育程度</asp:ListItem>
                                    <asp:ListItem Value="4">特定對象</asp:ListItem>
                                    <asp:ListItem Value="5">結訓後動向</asp:ListItem>
                                    <asp:ListItem Value="6">工作年資</asp:ListItem>
                                    <asp:ListItem Value="7">受訓學員地理分布</asp:ListItem>
                                    <asp:ListItem Value="8">所屬公司行業別</asp:ListItem>
                                    <asp:ListItem Value="9">所屬公司規模</asp:ListItem>
                                    <asp:ListItem Value="10">參訓動機</asp:ListItem>
                                    <asp:ListItem Value="15">參加單位類別</asp:ListItem>
                                    <asp:ListItem Value="16">職能課程分類</asp:ListItem>
                                    <asp:ListItem Value="17">參加課程型態</asp:ListItem>
                                    <asp:ListItem Value="18">外籍配偶類別</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">預算來源</td>
                            <td class="whitecol" id="tdBudgetList2" runat="server">
                                <span runat="server" id="spBudgetList1">
                                    <asp:CheckBoxList ID="BudgetList" runat="server" RepeatDirection="Horizontal">
                                    </asp:CheckBoxList>
                                    <asp:Label ID="Label2" runat="server" Text="Label" ForeColor="Red">(註：如未勾選則統計全部預算別,含未選擇預算別或不補助)</asp:Label>
                                </span>
                                <input id="BudgetHidden" type="hidden" value="0" runat="server" name="BudgetHidden" />
                                <span runat="server" id="spBudgetList2">
                                    <asp:Label ID="Label1" runat="server" Text="Label" ForeColor="Red">(無此選項)</asp:Label></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <p align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_Export_M"></asp:Button>
                    </p>
                    <table id="DataGroupTable" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td>
                                <div id="Div1" runat="server">
                                    <asp:Table ID="DataTable1" runat="server" Width="100%" CssClass="table_sch" BackColor="WhiteSmoke" CellSpacing="0" CellPadding="2"></asp:Table>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="BtnPrint2" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                &nbsp;
                                <asp:Button ID="BtnExp2" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
