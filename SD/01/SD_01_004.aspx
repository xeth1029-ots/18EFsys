<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_01_004.aspx.vb" Inherits="WDAIIP.SD_01_004" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>e網報名審核</title>
    <meta name="generator" content="microsoft visual studio .net 7.1" />
    <meta name="code_language" content="visual basic .net 7.1" />
    <meta name="vs_defaultclientscript" content="javascript" />
    <meta name="vs_targetschema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <%--<script type="text/javascript" src="../../js/date-picker2.js"></script>--%>
    <script type="text/javascript" src="../../js/date-picker2.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181018
        <%-- 
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);
        --%>
        //'查詢參訓歷史 'open_StudentList
        function open_StudentList(Button1ClientID, rqID) {
            var CST_KD_STUDENTLIST = 'StudentList';
            window.open('../05/SD_05_010_pop.aspx?ID=' + rqID + '&SD_01_004_Type=' + CST_KD_STUDENTLIST + '&BtnHistory=' + Button1ClientID + '', 'history', 'width=1400,height=820,scrollbars=1')
            return false;
        }

        //'預算別
        function Change(ddlBudID, ddlSupplyID) {
            var ddlBudIDobj = document.getElementById(ddlBudID);
            var ddlSupplyIDobj = document.getElementById(ddlSupplyID);
            if (!ddlBudIDobj || !ddlSupplyIDobj) { return; }
            //特定100%
            if (ddlBudIDobj.value == '97') { ddlSupplyIDobj.value = '2'; }
            //0%			
            if (ddlBudIDobj.value == '99') { ddlSupplyIDobj.value = '9'; }
        }

        function GETvalue() {
            document.getElementById('Button7').click();
        }

        //var cst_DataGrid1_0 = 0
        var cst_Name = 1; //姓名
        var cst_signUpStatus = 6; //報名審核-報名成功或失敗。
        //var cst_HBudID = 8; //預算別/補助比例。
        //班級名稱
        function choose_class() {
            var RID = document.form1.RIDValue.value;
            document.form1.TMID1.value = '';
            document.form1.TMIDValue1.value = '';
            document.form1.OCID1.value = '';
            document.form1.OCIDValue1.value = '';
            document.form1.hidLockTime1.value = '1'; //1:鎖定
            openClass('../02/SD_02_ch.aspx?special=11&RID=' + document.form1.RIDValue.value);
        }

        //obj: signUpStatus.ClientID
        function ChangeStatus(num, obj) {
            //num 1:ok 2:ng
            document.getElementById(obj).value = num;
        }

        //存檔檢核(產投檢核。)
        function CheckData(col) {
            var MyTable = document.getElementById('DataGrid1');
            var num = col;
            var msg = '';
            var noValue1 = "請選擇";
            var noValue2 = "";
            //num debugger;
            for (i = 1; i < MyTable.rows.length; i++) {
                //'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                var cells_signUpStatus = MyTable.rows[i].cells[cst_signUpStatus];

                if (cells_signUpStatus.children.length > 0) {
                    if (cells_signUpStatus.children[0].checked) {
                        //debugger;//hidstar3-重複參訓加強提示功能
                        if (cells_signUpStatus.children[3].value != '') {
                            document.form1.hidstar3.value = cells_signUpStatus.children[3].value;
                        }
                    }
                    //num debugger;
                    if (cells_signUpStatus.children[1].checked && MyTable.rows[i].cells[num].childNodes[1].children[0].value == "") {
                        msg += '請輸入審核失敗原因.(第' + i + '行:' + MyTable.rows[i].cells[cst_Name].innerHTML + ')\n';
                    }
                }
            }

            //hidstar3-重複參訓加強提示功能
            if (document.form1.hidstar3.value != '') {
                if (!confirm('本次e網審核通過之學員,仍有學員在訓中,是否儲存,請確認!')) { msg += '學員,仍在訓中\n'; }
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
            else {
                return Chk_Blacklist();
            }
        }

        //存檔檢核(非產投、TIMS檢核。)
        function ErrmsgShow(col) {
            var MyTable = document.getElementById('DataGrid1');
            var num = col; //cst_失敗原因位置
            var msg = '';
            var hidstar3 = document.getElementById('hidstar3');//hidstar3-重複參訓加強提示功能
            var Hid_PreUseLimited18a = document.getElementById('Hid_PreUseLimited18a'); //限定2018年職前計畫
            var HidIJCMsg = document.getElementById('HidIJCMsg');
            HidIJCMsg.value = ""; //配合：Hid_IJC
            var Hid_ZIP2MSG1 = document.getElementById('Hid_ZIP2MSG1'); //Hid_MSG1NAME(姓名)
            var Hid_TV1_MSG1S = document.getElementById('Hid_TV1_MSG1S'); //尼伯特颱風臺東地區受災者
            var Hid_ZIP2MSG2 = document.getElementById('Hid_ZIP2MSG2'); //Hid_MSG2NAME(姓名)
            var Hid_TV1_MSG2S = document.getElementById('Hid_TV1_MSG2S'); //尼伯特颱風受災者,為屏東地區或臺南市七股區民眾
            var Hid_ZIP2MSG3 = document.getElementById('Hid_ZIP2MSG3'); //Hid_MSG2NAME(姓名)
            var Hid_TV2_MSG3S = document.getElementById('Hid_TV2_MSG3S'); //梅姬颱風受災者
            var Hid_DIS2MSG = document.getElementById('Hid_DIS2MSG'); //Hid_DIS2MSG(姓名)
            var Hid_DIS2ALARMMSG1 = document.getElementById('Hid_DIS2ALARMMSG1'); //屬於重大災害受災地區範圍
            //var Hid_oTest = document.getElementById('Hid_oTest'); 
            for (var i = 1; i < MyTable.rows.length; i++) {
                var cells_signUpStatus = MyTable.rows[i].cells[cst_signUpStatus];
                if (cells_signUpStatus.children.length > 0) {
                    //TEXTAREA
                    try {
                        if (cells_signUpStatus.children[1].checked && MyTable.rows[i].cells[num].childNodes[1].children[0].value == "") {
                            msg += '請輸入審核失敗原因..(第' + i + '行:' + MyTable.rows[i].cells[1].innerHTML + ')\n';
                        }
                    }
                    catch (err) {
                    }
                    if (cells_signUpStatus.children[0].checked) {
                        //debugger;//審核成功-重複參訓加強提示功能 3: hstar3 //hidstar3-重複參訓加強提示功能
                        if (cells_signUpStatus.children[3].value != '') {
                            if (hidstar3.value == '') {
                                hidstar3.value = cells_signUpStatus.children[3].value;
                            }
                        }
                    }
                    if (cells_signUpStatus.children[0].checked) {
                        //4:Hid_IJC:因與委外實施基準條款有抵觸，請確認是否要同意此民眾的報名
                        if (cells_signUpStatus.children[4].value != '') {
                            if (HidIJCMsg.value != "") { HidIJCMsg.value += "\n"; }
                            HidIJCMsg.value += cells_signUpStatus.children[4].value;
                        }
                    }
                    if (cells_signUpStatus.children[0].checked) {
                        //6:Hid_MSG1NAME:'檢核是否有郵遞區號訊息。[姓名]
                        //7:Hid_MSGTYPEN:'1:臺東地區民眾 2:屏東地區或臺南市七股區民眾 3:梅姬颱風受災者
                        //8:Hid_MSGADIDN:'重大災害受災地區範圍 序號
                        if (cells_signUpStatus.children[6].value != '') {
                            if (cells_signUpStatus.children[7].value == '1') {
                                if (Hid_ZIP2MSG1.value != "") { Hid_ZIP2MSG1.value += ","; }
                                Hid_ZIP2MSG1.value += cells_signUpStatus.children[6].value;
                            }
                            if (cells_signUpStatus.children[7].value == '2') {
                                if (Hid_ZIP2MSG2.value != "") { Hid_ZIP2MSG2.value += ","; }
                                Hid_ZIP2MSG2.value += cells_signUpStatus.children[6].value;
                            }
                            if (cells_signUpStatus.children[7].value == '3') {
                                if (Hid_ZIP2MSG3.value != "") { Hid_ZIP2MSG3.value += ","; }
                                Hid_ZIP2MSG3.value += cells_signUpStatus.children[6].value;
                            }
                            if (cells_signUpStatus.children[8].value != '') {
                                if (Hid_DIS2MSG.value != "") { Hid_DIS2MSG.value += ","; }
                                Hid_DIS2MSG.value += cells_signUpStatus.children[6].value;
                            }
                        }
                    }
                }
            }
            if (Hid_PreUseLimited18a.value == "") {
                //非限定中的職前計畫要顯示！//hidstar3-重複參訓加強提示功能
                if (hidstar3.value != '') {
                    if (!confirm('本次e網審核通過之學員,仍有學員在訓中,是否儲存,請確認!')) msg += '學員,仍在訓中\n';
                }
            }
            if (msg != '') {
                alert(msg);
                return false; //異常直接中斷。
            } else {
                var rst2 = true; //正常再次檢核。
                if (rst2) rst2 = Chk_Blacklist(); //正常再次檢核。
                if (rst2) rst2 = Chk_IJClist(); //正常再次檢核。
                if (rst2) {
                    if (Hid_ZIP2MSG1.value != '') {
                        //尼伯特颱風臺東地區受災者(1)
                        rst2 = confirm(Hid_ZIP2MSG1.value + Hid_TV1_MSG1S.value);
                    }
                }
                if (rst2) {
                    if (Hid_ZIP2MSG2.value != '') {
                        //尼伯特颱風臺東地區受災者(2)
                        rst2 = confirm(Hid_ZIP2MSG2.value + Hid_TV1_MSG2S.value);
                    }
                }
                if (rst2) {
                    if (Hid_ZIP2MSG3.value != '') {
                        //梅姬颱風受災者(3)
                        rst2 = confirm(Hid_ZIP2MSG3.value + Hid_TV2_MSG3S.value);
                    }
                }
                if (rst2) {
                    if (Hid_DIS2MSG.value != '') {
                        //屬於重大災害受災地區範圍
                        rst2 = confirm(Hid_DIS2MSG.value + Hid_DIS2ALARMMSG1.value);
                    }
                }
                return rst2;
            }
        }

        //全選
        function SelectAll(num) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                //'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                var cells_signUpStatus = MyTable.rows[i].cells[cst_signUpStatus];
                if (cells_signUpStatus.children.length != 0 && !cells_signUpStatus.children[0].disabled && !cells_signUpStatus.children[1].disabled) {
                    cells_signUpStatus.children[0].checked = (num == 1) ? true : false;
                    cells_signUpStatus.children[1].checked = (num == 1) ? false : true;
                    cells_signUpStatus.children[2].value = num;
                }
            }
        }

        //黑名單
        function Chk_Blacklist() {
            //黑名單 //警告訊息，但確認後可繼續儲存。
            var rst1 = true;
            var msg = document.getElementById('hidBlackMsg').value;
            if (msg != "") {
                msg += "\n詳情請至教務管理-學員黑名單查詢\n是否續繼儲存?";
                if (!confirm(msg)) {
                    rst1 = false;
                }
            }
            return rst1;
        }

        function Chk_IJClist() {
            //因與委外實施基準條款有抵觸，請確認是否要同意此民眾的報名 //警告訊息，但確認後可繼續儲存。
            var rst1 = true;
            var HidIJCMsg = document.getElementById('HidIJCMsg');
            var msg = HidIJCMsg.value;
            if (msg != "") {
                if (!confirm(msg)) {
                    rst1 = false;
                }
            }
            return rst1;
        }

        /*個資法js*/
        function showLoginPwdDiv(num) {
            //num: 1:查詢 2:匯出 (記錄目前查詢按鈕) //return false;
            var hidSchBtnNum = document.getElementById('hidSchBtnNum'); //記錄目前查詢按鈕
            hidSchBtnNum.value = num; //num: 1:查詢 2:匯出 (記錄目前查詢按鈕)
            //var rblWorkMode_0 = document.getElementById('rblWorkMode_0');   //模糊顯示
            //var rblWorkMode_1 = document.getElementById('rblWorkMode_1');   //正常顯示
            var hidLockTime1 = document.getElementById('hidLockTime1');   //啟用鎖定 1:鎖定 2:不鎖定。
            var hidLockTime2 = document.getElementById('hidLockTime2');
            var OCIDValue1 = document.getElementById('OCIDValue1');
            var IDNO = document.getElementById('IDNO');
            var divPwdFrame = document.getElementById('divPwdFrame');
            var txtdivPxssward = document.getElementById('txtdivPxssward');
            //(hidLockTime1)啟用鎖定 1:鎖定 2:不鎖定。
            if (OCIDValue1.value == '' && IDNO.value == '') { hidLockTime1.value == '1'; }

            //if (rblWorkMode_1.checked == true) {
            //    if (OCIDValue1.value == '' && IDNO.value == '') {
            //        hidLockTime1.value == '1'; //啟用鎖定 1:鎖定 2:不鎖定。
            //    }
            //}
            var blnPwdFrame = false; //不顯示密碼輸入
            //if (rblWorkMode_1.checked != true) { hidLockTime1.value = '1'; }
            //if (rblWorkMode_1.checked == true && hidLockTime1.value == '1' && hidLockTime2.value == '1') {
            //    blnPwdFrame = true; //顯示密碼輸入
            //}
            //alert(hidLockTime1.value);
            if (blnPwdFrame) {
                divPwdFrame.style.display = 'inline'; //顯示
                if (txtdivPxssward != null) txtdivPxssward.focus();
                return false;
            }
            else {
                document.getElementById('divPwdFrame').style.display = 'none';
                return true;
            }
        }

        function chkTxtPassword() {
            //num: 1:查詢 2:匯出 (記錄目前查詢按鈕)
            var divPwdFrame = document.getElementById('divPwdFrame');
            var txtdivPxssward = document.getElementById('txtdivPxssward');
            var labChkMsg = document.getElementById('labChkMsg');
            var msg = '';
            if (txtdivPxssward.value == '') msg = '請輸入您的個資安全密碼!';
            //divPwdFrame.style.display = 'none';
            if (msg != '') {
                //debugger;
                labChkMsg.innerText = msg;
                alert(msg);
                return false;
            }
        }
    </script>
    <%--<style type="text/css"> #File1 { width: 300px; } </style>--%>
    <style type="text/css">
        .AAstyle1 { color: #000000; font-weight: bolder; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="position: absolute; top: -333px">
            <input type="text" title="Chaff for Chrome Smart Lock" /><input type="password" title="Chaff for Chrome Smart Lock" />
        </div>
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%" class="font">
            <tr>
                <td align="center">
                    <table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;e網報名審核</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="table3" class="table_nw" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%" Columns="45" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button8" value="..." type="button" runat="server" class="asp_button_Mini" />
                                <input id="DistValue" type="hidden" runat="server" />
                                <asp:Button Style="display: none" ID="Button7" runat="server" Text="Button7" class="asp_button_M"></asp:Button>
                                <span style="position: absolute; display: none" id="HistoryList2" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級名稱 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button5" onclick="choose_class();" value="..." type="button" runat="server" class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span style="position: absolute; display: none; left: 28%" id="HistoryList">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">通俗職類 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" Columns="30" onfocus="this.blur()" Width="40%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" value="..." type="button" runat="server" class="asp_button_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">身分證號碼 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="IDNO" runat="server" MaxLength="20" Width="30%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">報名日期 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="start_date" runat="server" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.clientid %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" />
                                </span>
                                ～
                                <asp:TextBox ID="end_date" runat="server" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.clientid %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" />
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">審核狀態 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="SsignUpStatus" runat="server" RepeatDirection="horizontal" RepeatLayout="flow" CssClass="font">
                                    <asp:ListItem Value="1" Selected="true">尚未審核</asp:ListItem>
                                    <asp:ListItem Value="2">審核成功</asp:ListItem>
                                    <asp:ListItem Value="3">審核失敗</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr id="TPLANID28_TR1" runat="server">
                            <td class="bluecol">匯入e網報名名冊<asp:Label ID="labPlanTxt" runat="server" Text="(產業人才投資)"></asp:Label>
                            </td>
                            <td class="whitecol">
                                <input id="File1" type="file" name="File1" runat="server" size="70" accept=".xls,.ods" />
                                <asp:Button ID="BtnImport28" runat="server" Text="匯入名冊" CssClass="asp_button_M" Enabled="False"></asp:Button>(必須為ods或xls格式)
                                <asp:HyperLink ID="HyperLink1" runat="server" NavigateUrl="../../doc/Stud_Temp2.zip" CssClass="font" ForeColor="#8080ff">下載整批上載格式檔</asp:HyperLink>
                                <asp:Button ID="Button11" runat="server" Text="列印匯入學員報名名冊用的班級代碼" CssClass="asp_Export_M"></asp:Button>
                                <asp:CheckBox ID="CheckBox1" runat="server" Text="檢視" ForeColor="silver" Visible="false"></asp:CheckBox>
                            </td>
                        </tr>
                        <%--<tr>
                            <td class="bluecol">資料顯示模式 </td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="rblWorkMode" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="1">模糊顯示</asp:ListItem>
                                    <asp:ListItem Value="2" Selected="True">正常顯示</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>--%>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td colspan="3" class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td colspan="3" class="whitecol">
                                <asp:DataGrid ID="dtgAddresses1" runat="server" CellPadding="8" GridLines="both" CssClass="font" Width="100%">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="slateblue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <%--
							    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S" OnClientClick="return showLoginPwdDiv(1);" CommandName="Button1"></asp:Button>&nbsp;
							    <asp:Button ID="Button13" runat="server" Text="匯出" CssClass="asp_button_S" OnClientClick="return showLoginPwdDiv(2);" CommandName="Button13"></asp:Button>
                                --%>
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M" CommandName="Button1"></asp:Button>&nbsp;
							    <asp:Button ID="Button13" runat="server" Text="匯出" CssClass="asp_Export_M" CommandName="Button13"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="red"></asp:Label>
                </td>
            </tr>
        </table>
        <table id="DataGridTable" class="font" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
            <tr>
                <td>&nbsp;&nbsp;<asp:Label ID="Label1" runat="server" ForeColor="green">* 表示該學員尚有產投或自辦在職課程仍在訓中，請查詢學員參訓歷史</asp:Label><br />
                    <%--職場續航-<asp:LinkButton ID="LinkButton1" runat="server" Visible="false">測試寄送信件</asp:LinkButton>--%>
                    &nbsp;&nbsp;<asp:Label ID="LabWYROLE" runat="server" ForeColor="Red">* 表示該學員符合「職場續航」優先錄訓條件，可將滑鼠移至姓名處查看年資、年齡。<br /></asp:Label>
                    &nbsp;&nbsp;<asp:Label ID="LabWYROLE2" runat="server">(優先錄訓條件：1-工作15年以上年滿55歲、2-工作25年以上、3-工作10年以上年滿60歲、4-強制退休年齡前2年內之63-64歲者)<br /></asp:Label>
                    &nbsp;&nbsp;<asp:Label ID="LabSubsidyCost" runat="server" ForeColor="blue">* 表示為該學員已申請職訓生活津貼,可點選檢視功能查詢<br /></asp:Label>
                    &nbsp;&nbsp;<asp:Label ID="Label2" runat="server" ForeColor="Blue">姓名藍色表該學員非報名本班同一天甄試日期有資料,序號藍色表示報名序號</asp:Label>
                    <div id="divEnterDouble" runat="server">&nbsp;&nbsp;<asp:Label ID="labEnterDouble" runat="server" CssClass="AAstyle1">報名時段重疊名單：</asp:Label></div>
                    <div id="divEnterMoney" runat="server">&nbsp;&nbsp;<asp:Label ID="labEnterMoney" runat="server" CssClass="AAstyle1">補助費已達6萬名單：</asp:Label></div>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" AllowPaging="true" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="序號">
                                <HeaderStyle Wrap="false" HorizontalAlign="Center" Width="4%" VerticalAlign="middle"></HeaderStyle>
                                <ItemStyle Wrap="false" HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="star2" runat="server" ForeColor="blue" CssClass="font">*</asp:Label>
                                    <asp:TextBox ID="stud1" runat="server" onfocus="this.blur()" Width="80%" ForeColor="red" MaxLength="4" Enabled="false"></asp:TextBox>
                                    <asp:Label ID="star3" runat="server" ForeColor="green" CssClass="font">*</asp:Label>
                                    <asp:Label ID="star4" runat="server" ForeColor="Red" CssClass="font">*</asp:Label>
                                </ItemTemplate>
                                <FooterStyle Wrap="false"></FooterStyle>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="姓名">
                                <HeaderStyle HorizontalAlign="Center" Width="6%" VerticalAlign="middle"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="labSTNAME" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <%--<asp:BoundColumn HeaderText="姓名" ItemStyle-HorizontalAlign="Center">
                                <HeaderStyle HorizontalAlign="Center" Width="6%" VerticalAlign="middle"></HeaderStyle>
                            </asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證號碼">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle" Width="8%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OrgName" HeaderText="報名機構">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle" Width="10%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn HeaderText="報名班級">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle" Width="10%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="RelEnterDate" HeaderText="報名日期" DataFormatString="{0:g}">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle" Width="8%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="報名審核">
                                <HeaderStyle Width="7%"></HeaderStyle>
                                <HeaderTemplate>
                                    報名審核<br>
                                    <input onclick="SelectAll(1);" type="radio" value="on" name="all">成功<input onclick="    SelectAll(2);" type="radio" name="all">失敗
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <%--cst_signUpStatus--%>
                                    <input id="signUpStatus1" type="radio" value="signUpStatus1" name="RadioGroup" runat="server" />成功
								    <input id="signUpStatus2" type="radio" value="signUpStatus2" name="RadioGroup" runat="server" />失敗
								    <input id="signUpStatus" type="hidden" name="signUpStatus" runat="server" />
                                    <input id="hstar3" type="hidden" name="hstar3" runat="server" />
                                    <input id="Hid_IJC" type="hidden" name="Hid_IJC" runat="server" />
                                    <input id="HidIdentityID" type="hidden" runat="server" />
                                    <input id="Hid_MSG1NAME" type="hidden" runat="server" />
                                    <input id="Hid_MSGTYPEN" type="hidden" runat="server" />
                                    <input id="Hid_MSGADIDN" type="hidden" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="報名路徑">
                                <HeaderStyle HorizontalAlign="Center" Width="8%" VerticalAlign="middle"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="middle"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Label ID="LabEnterPath" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <%--
                            <asp:TemplateColumn HeaderText="預算別">
							    <HeaderStyle Width="150px" Wrap="False"></HeaderStyle>
							    <HeaderTemplate>
								    預算別&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 補助比例<br>
								    <asp:DropDownList ID="HBudID" runat="server" Width="65px" Height="20px">
									    <asp:ListItem Value="01">公務</asp:ListItem>
									    <asp:ListItem Value="02">就安</asp:ListItem>
									    <asp:ListItem Value="03">就保</asp:ListItem>
									    <asp:ListItem Value="97">協助</asp:ListItem>
									    <asp:ListItem Value="99">不補助</asp:ListItem>
									    <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
								    </asp:DropDownList>
								    <select style="width: 85px" onchange="SelectCom();" name="HSupplyID">
									    <option value="1">
									    一般80%<option value="2">
									    特定100%<option value="9">
									    0%<option selected="selected">請選擇</option>
								    </select>
							    </HeaderTemplate>
							    <ItemTemplate>
								    <asp:DropDownList ID="BudID" runat="server" Width="65px" Height="20px">
									    <asp:ListItem Value="01">公務</asp:ListItem>
									    <asp:ListItem Value="02">就安</asp:ListItem>
									    <asp:ListItem Value="03">就保</asp:ListItem>
									    <asp:ListItem Value="97">協助</asp:ListItem>
									    <asp:ListItem Value="99">不補助</asp:ListItem>
									    <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
								    </asp:DropDownList>
								    <asp:DropDownList ID="SupplyID" runat="server" Width="82px">
									    <asp:ListItem Value="1">一般80%</asp:ListItem>
									    <asp:ListItem Value="2">特定100%</asp:ListItem>
									    <asp:ListItem Value="9">0%</asp:ListItem>
									    <asp:ListItem>請選擇</asp:ListItem>
								    </asp:DropDownList>
							    </ItemTemplate>
						    </asp:TemplateColumn>
                            --%>
                            <asp:TemplateColumn HeaderText="是否為在職者&lt;br&gt;補助身分">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle"></HeaderStyle>
                                <ItemTemplate>
                                    <input id="WorkSuppIdent1" type="radio" value="Y" name="WorkSuppIdentGroup" runat="server" />是
								<input id="WorkSuppIdent2" type="radio" value="N" name="WorkSuppIdentGroup" runat="server" />否
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <%--<asp:TemplateColumn HeaderText="協助基金">,<HeaderStyle Width="5%" />,<ItemStyle HorizontalAlign="Center" />,<ItemTemplate>
                               ,<asp:Label ID="BudgetID97" runat="server"></asp:Label>,</ItemTemplate>,</asp:TemplateColumn>--%>
                            <asp:BoundColumn DataField="actno" HeaderText="保險證號">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle" Width="7%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="備註(失敗原因)">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle" Width="8%"></HeaderStyle>
                                <ItemTemplate>
                                    <div>
                                        <asp:TextBox ID="signUpMemo" runat="server" MaxLength="150" TextMode="MultiLine" Width="95%" Height="60px"></asp:TextBox><br>
                                        <asp:Button ID="BtnHistory" runat="server" Text="近二年參訓歷史" CommandName="History" CssClass="asp_button_M"></asp:Button>
                                        <asp:Label ID="labDiffYears" runat="server"></asp:Label>
                                        <%--BtnHistory:報名及補助查詢 <input id="BtnHistory" runat="server" type ="button"  commandname="History" width="100px" />--%>
                                    </div>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="middle" Width="7%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" VerticalAlign="middle" Font-Size="Small"></ItemStyle>
                                <ItemTemplate>
                                    <asp:LinkButton ID="Button2" runat="server" Text="檢視" ToolTip="檢視學員報名資料" CommandName="view" CssClass="linkbutton"></asp:LinkButton>
                                    <asp:LinkButton ID="Button4" runat="server" Text="刪除" ToolTip="刪除學員報名資料" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
                                    <asp:LinkButton ID="Button6" runat="server" Text="還原" ToolTip="將報名資料還原至收件狀態" CommandName="rev" CssClass="linkbutton"></asp:LinkButton>
                                    <input id="HidBirthDay" type="hidden" runat="server" />
                                    <input id="HidSTDate" type="hidden" runat="server" />
                                    <input id="Hid_eSerNum" type="hidden" runat="server" />
                                    <input id="HidCMASTER1" type="hidden" runat="server" />
                                    <input id="HidCMASTER1NT" type="hidden" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="false"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="Button9" runat="server" Text="查詢參訓歷史" CssClass="asp_button_M"></asp:Button>&nbsp;
                    <asp:Button ID="Button3" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <input id="Years" type="hidden" name="Years" runat="server" />
                    <input id="hidBlackMsg" type="hidden" runat="server" />
                </td>
            </tr>
        </table>
        <div id="divPwdFrame" runat="server" style="position: absolute; border-width: 6px; border-style: double; border-color: #4682B4; display: none; width: 350px; height: 300px; left: 195px; top: 200px; background-color: #FFFAF0; padding-left: 30px; padding-top: 30px;">
            <table align="center">
                <tr>
                    <td align="center">請輸入個資安全密碼 </td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:TextBox ID="txtdivPxssward" runat="server" TextMode="Password"></asp:TextBox></td>
                </tr>
                <tr>
                    <td align="center"></td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Button ID="btndivPwdSubmit" runat="server" Text="確定" OnClientClick="return chkTxtPassword();" CssClass="asp_button_S" CommandName="btndivPwdSubmit" />
                        &nbsp;<input id="btn_close" type="button" value="關閉" onclick="document.getElementById('divPwdFrame').style.display = 'none'; document.getElementById('labChkMsg').text = '';" class="button_b_S" />
                    </td>
                </tr>
                <tr>
                    <td align="center"></td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Label ID="labChkMsg" runat="server" CssClass="needFont"></asp:Label></td>
                </tr>
            </table>
        </div>
        <input id="fileSizeLimit" type="hidden" value="4" runat="server" />
        <input id="hidstar3" type="hidden" runat="server" />
        <input id="hidLockTime1" type="hidden" runat="server" value="1" />
        <input id="hidSchBtnNum" type="hidden" runat="server" value="1" />
        <input id="isBlack" type="hidden" runat="server">
        <input id="Blackorgname" type="hidden" runat="server">
        <input id="HidIJCMsg" type="hidden" runat="server">
        <input id="hidLockTime2" type="hidden" runat="server" value="1" />
        <%--<input id="HidCIShow" type="hidden" name="HidCIShow" runat="server">--%>
        <asp:HiddenField ID="Hid_ZIP2MSG1" runat="server" />
        <asp:HiddenField ID="Hid_TV1_MSG1S" runat="server" />
        <asp:HiddenField ID="Hid_ZIP2MSG2" runat="server" />
        <asp:HiddenField ID="Hid_TV1_MSG2S" runat="server" />
        <asp:HiddenField ID="Hid_ZIP2MSG3" runat="server" />
        <asp:HiddenField ID="Hid_TV2_MSG3S" runat="server" />
        <asp:HiddenField ID="Hid_DIS2MSG" runat="server" />
        <asp:HiddenField ID="Hid_DIS2ALARMMSG1" runat="server" />
        <%--<asp:HiddenField ID="Hid_oTest" runat="server" />--%>
        <asp:HiddenField ID="Hid_PreUseLimited18a" runat="server" />
        <asp:HiddenField ID="Hid_impOCID" runat="server" />
    </form>
</body>
</html>
