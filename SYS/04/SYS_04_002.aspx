<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_002.aspx.vb" Inherits="WDAIIP.SYS_04_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>鍵詞維護</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function check_search() {
            if (document.FDUpdate.KeyType.selectedIndex == 0) {
                alert('請選擇鍵詞種類!!');
                return false;
            }
        }

        function returnValue17(kbsid, kbid, kbname, kbdesc1,
            mustfill, orgkindgw, ksort, uselatestver, downloadrpt, rptname, uploadfl1, sentbatver, usememo1) {
            //,datagrid08
            document.getElementById('hid_BIDCASE_KBSID').value = kbsid;
            document.FDUpdate.keycode.value = kbid;
            document.FDUpdate.keyname.value = kbname;
            document.getElementById('txt_KBDESC1').value = kbdesc1;
            document.getElementById('txt_RPTNAME').value = rptname;

            document.FDUpdate.CB_MUSTFILL.checked = (mustfill == "Y");
            for (var i = 0; i < document.FDUpdate.RBL_ORGKINDGW.length; i++) {
                document.FDUpdate.RBL_ORGKINDGW[i].checked = (document.FDUpdate.RBL_ORGKINDGW[i].value == orgkindgw) ? true : false;
            }
            document.getElementById('txt_KSORT').value = ksort;
            document.FDUpdate.cb_USELATESTVER.checked = (uselatestver == "Y");
            document.FDUpdate.cb_DOWNLOADRPT.checked = (downloadrpt == "Y");
            document.FDUpdate.cb_UPLOADFL1.checked = (uploadfl1 == "Y");
            document.FDUpdate.cb_SENTBATVER.checked = (sentbatver == "Y");
            document.FDUpdate.cb_USEMEMO1.checked = (usememo1 == "Y");
            //document.FDUpdate.cb_DataGrid08.checked = (datagrid08 == "Y");
            document.FDUpdate.keycode.disabled = true;
            document.FDUpdate.keyname.disabled = true;
        }

        function returnValue16(v_sbsid, v_swdate, v_monthwage, v_hourwage) {
            //實施日期:SWDATE /每月基本工資:MONTHLYWAGE /每小時基本工資:HOURLYWAGE
            document.getElementById('lab_BASICSALARY_SBSID').innerHTML = v_sbsid;
            document.getElementById('hid_BASICSALARY_SBSID').value = v_sbsid;
            document.getElementById('SWDATE').value = v_swdate;
            document.getElementById('MONTHLYWAGE').value = v_monthwage;
            document.getElementById('HOURLYWAGE').value = v_hourwage;
        }

        function returnValue15(KID, NAME, MEMO1) {
            document.FDUpdate.keycode.value = KID;
            document.FDUpdate.keyname.value = NAME;
            document.FDUpdate.txtmemo1.value = MEMO1;
        }

        function returnValue11(KID, NAME, Sort, ItemageName, ItemCostName) {
            document.FDUpdate.keycode.value = KID;
            document.FDUpdate.keyname.value = NAME;
            document.FDUpdate.Sort.value = Sort;
            document.FDUpdate.txtItemageName.value = ItemageName;
            document.FDUpdate.txtItemCostName.value = ItemCostName;
        }

        function returnValue40(KID, NAME, KeyTable) {
            document.FDUpdate.keycode.value = KID;
            document.FDUpdate.keyname.value = NAME;
            document.FDUpdate.txtKeyTable.value = KeyTable;
        }

        function set_IdentityID_123(value1) {
            //var str = ""; //debugger;
            var ilen = $("[id*='IdentityID'][type=checkbox]").length;
            if (ilen > 0) {
                $("[id*='IdentityID'][type=checkbox]").each(function (idx) {
                    var idx2 = parseInt($(this).attr("id").replace("IdentityID_", ""), 10);
                    if (isNaN(idx2)) { alert('isNaN'); reture; }
                    if (value1.substr(idx2, 1) == "1") {
                        $(this).prop("checked", true);//$(this).attr('checked');
                    }
                    else {
                        $(this).prop("checked", false);//$(this).removeAttr('checked');
                    }
                    //if (str != "") { str += ","; } str += idx2.toString() + "/" + value1.substr(idx2, 1);
                });
                //alert('ilen:' + ilen + ',str:' + str);
            }
        }

        //set_chkboxValue('cblSBLACK2TPLANID')
        function set_chkboxValue(value1, cblidname) {
            //var str = ""; //debugger; //cblSBLACK2TPLANID
            var ilen = $("[id*='" + cblidname + "'][type=checkbox]").length;
            if (ilen > 0) {
                $("[id*='" + cblidname + "'][type=checkbox]").each(function (idx) {
                    var idx2 = parseInt($(this).attr("id").replace(cblidname + "_", ""), 10);
                    if (isNaN(idx2)) { alert('isNaN'); reture; }
                    var v_checked = (value1.substr(idx2, 1) == "1") ? true : false;
                    $(this).prop("checked", v_checked); //$(this).attr('checked');
                });
                //alert('ilen:' + ilen + ',str:' + str);
            }
        }

        //KEY_PLAN
        function returnValue8(KID, NAME, PlanType, IsOnLine, ClsYear, QID, EmailSend, BlackList,
            QueryDisplay, IdentityValue, Reusable, useECFA, PropertyID, sblacktype, sbk2tplanid) {
            //debugger;
            document.FDUpdate.keycode.value = KID;
            document.FDUpdate.keyname.value = NAME;

            //IdentityID debugger;
            //setRadioValue(IdentityID, IdentityValue);
            //setCheckBoxList1(IdentityID, IdentityValue);
            set_IdentityID_123(IdentityValue);

            //radiobuttonlist
            for (var i = 0; i < document.FDUpdate.IsOnLine.length; i++) {
                document.FDUpdate.IsOnLine[i].checked = (document.FDUpdate.IsOnLine[i].value == IsOnLine) ? true : false;
            }
            document.FDUpdate.ClsYear.value = ClsYear;

            //dropdownlist
            if (PlanType == '') {
                document.FDUpdate.PlanType.selectedIndex = 0;
            }
            else {
                document.FDUpdate.PlanType.value = PlanType;
            }

            document.FDUpdate.TypeValue.value = PlanType;

            //dropdownlist
            if (PropertyID == '') {
                document.FDUpdate.ddlPropertyID.selectedIndex = 0;
            }
            else {
                document.FDUpdate.ddlPropertyID.value = PropertyID;
            }
            document.FDUpdate.HidPropertyID.value = PropertyID;

            //radiobuttonlist
            for (var i = 0; i < document.FDUpdate.rblEmailSend.length; i++) {
                document.FDUpdate.rblEmailSend[i].checked = (document.FDUpdate.rblEmailSend[i].value == EmailSend) ? true : false;
            }

            //radiobuttonlist
            for (var i = 0; i < document.FDUpdate.rblReusable.length; i++) {
                document.FDUpdate.rblReusable[i].checked = (document.FDUpdate.rblReusable[i].value == Reusable) ? true : false;
            }

            //radiobuttonlist
            for (var i = 0; i < document.FDUpdate.rblBlackList.length; i++) {
                document.FDUpdate.rblBlackList[i].checked = (document.FDUpdate.rblBlackList[i].value == BlackList) ? true : false;
            }
            //RadioButtonList sblacktype
            for (var i = 0; i < document.FDUpdate.rblSBLACKTYPE.length; i++) {
                document.FDUpdate.rblSBLACKTYPE[i].checked = (document.FDUpdate.rblSBLACKTYPE[i].value == sblacktype) ? true : false;
            }

            set_chkboxValue(sbk2tplanid, 'cblSBK2TPLANID')

            //radiobuttonlist
            for (var i = 0; i < document.FDUpdate.rbluseECFA.length; i++) {
                document.FDUpdate.rbluseECFA[i].checked = (document.FDUpdate.rbluseECFA[i].value == useECFA) ? true : false;
            }

            //radiobuttonlist
            for (var i = 0; i < document.FDUpdate.rblQueryDisplay.length; i++) {
                document.FDUpdate.rblQueryDisplay[i].checked = (document.FDUpdate.rblQueryDisplay[i].value == QueryDisplay) ? true : false;
            }

            //dropdownlist
            if (document.FDUpdate.KeyType.value == 'Key_Plan') {
                document.FDUpdate.QuesType.value = QID;
                if (QID == '') { alert('請選擇問卷類別!!'); }
            }

            ChangerblReusable();
        }

        function returnValue13(KID, NAME, UnUsedYear, MergeID, Subsidy, vSORT28, vSUPPLYID, vNOSHOWMI) {
            document.FDUpdate.keycode.value = KID;
            document.FDUpdate.keyname.value = NAME;
            //dropdownlist
            document.FDUpdate.chkSubsidy.checked = Subsidy;
            document.FDUpdate.ddlUnUsedYear.value = UnUsedYear;
            document.FDUpdate.ddlMergeID.value = MergeID;
            document.FDUpdate.txtSORT28.value = vSORT28;
            //radiobuttonlist 補助比例
            for (var i = 0; i < document.FDUpdate.rblSUPPLYID.length; i++) {
                document.FDUpdate.rblSUPPLYID[i].checked = (document.FDUpdate.rblSUPPLYID[i].value == vSUPPLYID) ? true : false;
            }
            //radiobuttonlist 主要參訓身分別不顯示
            for (var i = 0; i < document.FDUpdate.rblNOSHOWMI.length; i++) {
                document.FDUpdate.rblNOSHOWMI[i].checked = (document.FDUpdate.rblNOSHOWMI[i].value == vNOSHOWMI) ? true : false;
            }
            ChangeddlMergeID();
        }

        //KEY_LEAVE
        function returnValueKL(KID, NAME, MinusPoint, nouse, leavesort, engname) {
            document.FDUpdate.keycode.value = KID;
            document.FDUpdate.keyname.value = NAME;
            document.FDUpdate.MinusPoint.value = MinusPoint;
            document.FDUpdate.cb_LEAVE_NOUSE.checked = (nouse == 'Y') ? true : false;
            document.FDUpdate.Sort.value = leavesort;
            document.FDUpdate.EngkeyName.value = engname;
        }

        function returnValue(KID, NAME, nn1, nn2, MinusPoint, AddMinus, point, Type, online, Sort, DGHour, ClsYear, Type2, DegreeType) {
            document.FDUpdate.keycode.value = KID;
            document.FDUpdate.keyname.value = NAME;
            document.FDUpdate.Levels.value = nn1;
            document.FDUpdate.Parent1.value = nn2;
            document.FDUpdate.MinusPoint.value = MinusPoint;
            document.FDUpdate.AddMinus.value = AddMinus;
            document.FDUpdate.point.value = point;
            //radiobuttonlist 適用對象
            for (var i = 0; i < document.FDUpdate.rblDegreeType.length; i++) {
                document.FDUpdate.rblDegreeType[i].checked = false;
                if (document.FDUpdate.rblDegreeType[i].value == DegreeType) {
                    document.FDUpdate.rblDegreeType[i].checked = true;
                }
            }

            document.FDUpdate.Sort.value = Sort;
            document.FDUpdate.DGHour.value = DGHour;
        }

        function display_Item() {
            var BusID = document.FDUpdate.BusID;
            var JobID = document.FDUpdate.JobID;

            BusID.style.display = 'none';
            JobID.style.display = 'none';

            var mydrop = document.getElementById('KeyType');
            if (mydrop.value == 'Key_TrainType2') {
                BusID.style.display = '';
            }
            if (mydrop.value == 'Key_TrainType3') {
                Train(1);
                BusID.style.display = '';
                JobID.style.display = '';
            }
        }

        function Get_TMID() {
            if (document.FDUpdate.JobID.selectedIndex != 0) {
                var mydrop3 = document.getElementById('JobID');
                document.FDUpdate.TMIDValue.value = mydrop3.value;
            }
        }

        function ChangerblReusable() {
            //debugger;
            var cst_pt1 = 0;
            var cst_pt2 = 1;
            if (document.getElementsByName('rblReusable').length > 2) {
                cst_pt1 = 1; //cst_pt
                cst_pt2 = 2;
            }

            if (getRadioValue(document.FDUpdate.rblEmailSend) == 'N') {
                document.getElementsByName('rblReusable')[cst_pt1].checked = false;
                document.getElementsByName('rblReusable')[cst_pt2].checked = true;
                document.getElementsByName('rblReusable')[cst_pt1].disabled = true;
                document.getElementsByName('rblReusable')[cst_pt2].disabled = true;
            }
            else {
                document.getElementsByName('rblReusable')[cst_pt1].disabled = false;
                document.getElementsByName('rblReusable')[cst_pt2].disabled = false;
            }
        }

        //某些功能鎖定!
        function ChangeddlMergeID() {
            var v_ddlUnUsedYear = getValue("ddlUnUsedYear");
            if (v_ddlUnUsedYear == '') { document.FDUpdate.ddlMergeID.value = ""; }
            document.FDUpdate.ddlMergeID.disabled = (v_ddlUnUsedYear == '') ? true : true;
            //document.FDUpdate.ddlMergeID.style.display = (v_ddlUnUsedYear == '') ? 'none' : 'inline';
        }

        function check_save() {
            var msg = '';
            if (Table13.style.display == '') {
                if (getValue("ddlUnUsedYear") != '' && getValue("ddlMergeID") == '') {
                    //msg+='請選擇併入身分別!!\n';
                    //2010-02-03同意可不選擇併入身分別
                }
                if (getValue("ddlMergeID") != '') {
                    if (getValue("ddlMergeID") == document.FDUpdate.keycode.value) {
                        msg += '併入身分別與鍵值代碼相同，有誤請重新選擇!!\n';
                    }
                }
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        //選擇全部，若有單選消除全部勾選
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
    <%--<style type="text/css">
        .auto-style1 { color: #FF0000; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 45px; }
        .auto-style2 { color: #333333; padding: 4px; height: 45px; }
        .auto-style3 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 45px; }
    </style>--%>
    <style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; width: 20%; height: 78px; }
        .auto-style2 { color: #333333; padding: 4px; width: 80%; height: 78px; }
    </style>
</head>
<body>
    <form id="FDUpdate" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;鍵詞維護</asp:Label>
                </td>
            </tr>
        </table>
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td class="bluecol" style="width: 20%">鍵詞種類</td>
                <td class="whitecol">
                    <asp:DropDownList ID="KeyType" runat="server">
                        <asp:ListItem Value="===請選擇===">===請選擇===</asp:ListItem>
                    </asp:DropDownList>
                    <asp:DropDownList ID="BusID" runat="server"></asp:DropDownList>
                    <asp:DropDownList ID="JobID" runat="server"></asp:DropDownList>
                    &nbsp;<asp:Button ID="btnSearch" runat="server" Text="查詢" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    &nbsp;<asp:Button ID="btnConfigReset" runat="server" Text="系統參數重新載入" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    <input id="Levels" style="width: 6%" type="hidden" size="2" name="Levels" runat="server">
                    <input id="Parent1" style="width: 6%" type="hidden" name="Parent1" runat="server">
                    <input id="TMIDValue" style="width: 6%" type="hidden" size="3" name="Hidden1" runat="server">
                </td>
            </tr>
        </table>
        <br />
        <table class="table_nw" id="table1" width="100%">
            <tr>
                <td valign="top" width="24%" class="whitecol">
                    <%--
                <div id="keyUnit" style="border-bottom: #000000 1px solid; border-left: #000000 1px solid; background-color: #ffffcc; width: 250px; font-family: 新細明體; height: 300px; color: #000000; font-size: 9px; overflow: scroll; border-top: #000000 1px solid; font-weight: normal; border-right: #000000 1px solid" runat="server">
                    <iewc:treeview id="tvUnit" runat="server"></iewc:treeview>
                </div>
                    --%>
                    <div id="keyUnit" runat="server" style="border-bottom: #000000 1px solid; border-left: #000000 1px solid; background-color: #ffffcc; width: 100%; font-family: 新細明體; height: 300px; color: #000000; font-size: 9px; overflow: scroll; border-top: #000000 1px solid; font-weight: normal; border-right: #000000 1px solid">
                        <asp:TreeView ID="tvUnit" runat="server" ForeColor="Black"></asp:TreeView>
                    </div>
                </td>
                <td valign="top" class="whitecol">
                    <table cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td colspan="4">
                                <table id="Table2" runat="server" width="100%" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">扣分</td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:TextBox ID="MinusPoint" runat="server" Width="55%"></asp:TextBox></td>
                                        <td class="bluecol" style="width: 20%">停用</td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:CheckBox ID="cb_LEAVE_NOUSE" runat="server" Text="停用" /></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">英文欄位名稱</td>
                                        <td class="whitecol" style="width: 80%" colspan="3">
                                            <asp:TextBox ID="EngkeyName" runat="server" Width="55%"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table id="Table4" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">行業別</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:Label ID="Bus" runat="server"></asp:Label></td>
                                    </tr>
                                </table>
                                <table id="Table5" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">職業分類</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:Label ID="Job" runat="server"></asp:Label></td>
                                    </tr>
                                </table>
                                <table id="Table6" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">加扣</td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:TextBox ID="AddMinus" runat="server" Width="60%"></asp:TextBox></td>
                                        <td class="bluecol_need" style="width: 20%">分數</td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:TextBox ID="point" runat="server" Width="60%"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table id="Table8" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">可否線上報名</td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:RadioButtonList ID="IsOnLine" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="bluecol_need" style="width: 20%">停用年度</td>
                                        <td class="whitecol" style="width: 30%">
                                            <asp:TextBox ID="ClsYear" runat="server" Width="60%" MaxLength="4"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">e網報名審核發送Email</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="rblEmailSend" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <%--<td class="bluecol_need">是否下放轄區中心決定</td>--%>
                                        <td class="bluecol_need">是否下放轄區分署決定</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="rblReusable" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">是否適用於<b>ECFA</b>協助基金</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButtonList ID="rbluseECFA" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <%--SBLACKTYPE/'0:未設定/'1：各計畫自行限制處分紀錄/'2：跨計畫合併限制處分紀錄，因跨計畫合併限制處分紀錄可能會有不同組合，需要另外一個欄位紀錄組合喔/'3：所有計畫合併限制處分紀錄/'4：無處分限制(停用處分)--%>
                                        <td class="bluecol_need">學員處分功能</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="rblSBLACKTYPE" runat="server" RepeatLayout="Flow">
                                                <asp:ListItem Value="0">未設定</asp:ListItem>
                                                <asp:ListItem Value="1">各計畫自行限制處分紀錄</asp:ListItem>
                                                <asp:ListItem Value="2">跨計畫合併限制處分紀錄</asp:ListItem>
                                                <asp:ListItem Value="3">所有計畫合併限制處分紀錄</asp:ListItem>
                                                <asp:ListItem Value="4">無處分限制(停用處分)</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="bluecol">學員處分功能選擇<br />
                                            -<span style="color: red">跨計畫合併限制處分</span><br />
                                            -計畫必需多選</td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="cblSBK2TPLANID" runat="server" RepeatColumns="1" RepeatLayout="Flow"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">訓練單位處分功能</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButtonList ID="rblBlackList" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="Y">啟用</asp:ListItem>
                                                <asp:ListItem Value="N">停用</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">查詢時是否顯示</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButtonList ID="rblQueryDisplay" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="Y">是</asp:ListItem>
                                                <asp:ListItem Value="N">否</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">計畫分類</td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="PlanType" runat="server">
                                                <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                                <asp:ListItem Value="1">自辦</asp:ListItem>
                                                <asp:ListItem Value="2">委辦</asp:ListItem>
                                                <asp:ListItem Value="3">合辦</asp:ListItem>
                                                <asp:ListItem Value="4">補助</asp:ListItem>
                                            </asp:DropDownList>
                                            <input id="TypeValue" type="hidden" size="9" name="TypeValue" runat="server">
                                        </td>
                                        <td class="bluecol">問卷類別<font color="#ffffff">&nbsp;&nbsp;</font></td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="QuesType" runat="server"></asp:DropDownList></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">訓練性質</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:DropDownList ID="ddlPropertyID" runat="server">
                                                <asp:ListItem Value="N">停用</asp:ListItem>
                                                <asp:ListItem Value="0">職前</asp:ListItem>
                                                <asp:ListItem Value="1">在職</asp:ListItem>
                                            </asp:DropDownList>
                                            <input id="HidPropertyID" type="hidden" size="9" name="HidPropertyID" runat="server">
                                        </td>
                                    </tr>
                                </table>
                                <table id="Table7" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">排序</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="Sort" runat="server" Width="20%"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table id="Table9" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">單元時數</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:TextBox ID="DGHour" runat="server" Width="20%"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table id="Table10" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">適用對象</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:RadioButtonList ID="rblDegreeType" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="0">不拘</asp:ListItem>
                                                <asp:ListItem Value="1">個人</asp:ListItem>
                                                <asp:ListItem Value="2">班級</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                </table>
                                <table id="Table3" width="100%" cellpadding="1" cellspacing="1" runat="server">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">鍵值代碼</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:TextBox ID="keycode" runat="server" Width="66%"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">鍵值名稱</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="keyname" runat="server" Columns="30" Width="88%"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table id="Table15" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">鍵值說明</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:TextBox ID="txtmemo1" runat="server" Columns="60" Width="88%"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table id="Table16" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">資料序號</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:Label ID="lab_BASICSALARY_SBSID" runat="server" Text=""></asp:Label><asp:HiddenField ID="hid_BASICSALARY_SBSID" runat="server" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">實施日期</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:TextBox ID="SWDATE" runat="server" Width="48%" MaxLength="10"></asp:TextBox>(日期格式：yyyy/MM/dd)</td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">每月基本工資</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:TextBox ID="MONTHLYWAGE" runat="server" Width="48%" MaxLength="10"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">每小時基本工資</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:TextBox ID="HOURLYWAGE" runat="server" Width="48%" MaxLength="10"></asp:TextBox></td>
                                    </tr>
                                </table>

                                <table id="Table11" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">數量名稱</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:TextBox ID="txtItemageName" runat="server" Width="30%" ToolTip="不填寫將存為NULL"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table id="Table11b" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">計價單位</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:TextBox ID="txtItemCostName" runat="server" Width="20%" ToolTip="不填寫將存為NULL"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table id="Table12" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol_need" style="width: 20%">KeyTable</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:TextBox ID="txtKeyTable" runat="server" Width="30%"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <table id="Table13" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">生活津貼</td>
                                        <td class="whitecol" style="width: 80%">
                                            <asp:CheckBox Style="z-index: 0" ID="chkSubsidy" runat="server" Text="可申請生活津貼"></asp:CheckBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">停用年度</td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="ddlUnUsedYear" runat="server"></asp:DropDownList></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">併入身分別 <%--<P align="center"><font color="#ffffff">&nbsp;併入身分別</font></P>--%>
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="ddlMergeID" runat="server"></asp:DropDownList></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">排序序號</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txtSORT28" runat="server" Width="30%" MaxLength="10"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">補助比例</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="rblSUPPLYID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">一般80%</asp:ListItem>
                                                <asp:ListItem Value="2">特定100%</asp:ListItem>
                                                <asp:ListItem Value="9">0%</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">主要參訓身分別不顯示</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="rblNOSHOWMI" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="Y">不顯示</asp:ListItem>
                                                <asp:ListItem Value="N">顯示</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                </table>
                                <table id="Table14" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="bluecol" style="width: 20%">可使用身分別</td>
                                        <td class="whitecol" style="width: 80%">
                                            <input id="IdentityIDHidden" type="hidden" value="0" name="IdentityIDHidden" runat="server">
                                            <asp:CheckBoxList ID="IdentityID" runat="server" RepeatColumns="3"></asp:CheckBoxList>
                                        </td>
                                    </tr>
                                </table>
                                <table id="Table17" width="100%" runat="server" cellpadding="1" cellspacing="1">
                                    <tr>
                                        <td class="auto-style1">線上申辦項目</td>
                                        <td class="auto-style2">
                                            <asp:CheckBox ID="CB_MUSTFILL" runat="server" Text="必填資訊／非(免附文件)"></asp:CheckBox>
                                            <asp:CheckBox ID="cb_USELATESTVER" runat="server" Text="以最近一次版本送件"></asp:CheckBox>
                                            <asp:CheckBox ID="cb_DOWNLOADRPT" runat="server" Text="可下載報表"></asp:CheckBox>
                                            <br />
                                            <asp:CheckBox ID="cb_UPLOADFL1" runat="server" Text="檔案上傳"></asp:CheckBox>
                                            <asp:CheckBox ID="cb_SENTBATVER" runat="server" Text="以目前版本批次送出"></asp:CheckBox>
                                            <asp:CheckBox ID="cb_USEMEMO1" runat="server" Text="備註說明"></asp:CheckBox>
                                            <br />
                                            <%--<asp:CheckBox ID="cb_DataGrid08" runat="server" Text="訓練班別計畫表"></asp:CheckBox>--%>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">計畫類別</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="RBL_ORGKINDGW" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="G">產投</asp:ListItem>
                                                <asp:ListItem Value="W">自主</asp:ListItem>
                                            </asp:RadioButtonList>
                                            <asp:HiddenField ID="hid_BIDCASE_KBSID" runat="server" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">排序序號</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txt_KSORT" runat="server" Width="30%" MaxLength="10"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">報表名稱按鈕</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txt_RPTNAME" runat="server" Width="70%" MaxLength="100"></asp:TextBox></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">文字說明</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="txt_KBDESC1" runat="server" Width="70%" MaxLength="2000" Rows="7" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                </table>
                                <%--<asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ValidationExpression="[0-9]+" ControlToValidate="keycode" Display="None" ErrorMessage="請輸入數字"></asp:RegularExpressionValidator>--%>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4" align="center">
                                <asp:Button ID="cmdUpdate" runat="server" Text="修改" CssClass="asp_button_M"></asp:Button>&nbsp;
							<asp:Button ID="cmdAppend" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <%--<asp:ValidationSummary ID="ValidationSummary1" runat="server" DisplayMode="List" ShowMessageBox="True" ShowSummary="False"></asp:ValidationSummary>--%>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
