<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CP_01_006_add8.aspx.vb" Inherits="WDAIIP.CP_01_006_add8" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>不預告實地訪查紀錄表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <%--<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = false;
        if (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) { _isIE = true; }

        //決定date-picker元件使用的是西元年or民國年，by:20181018
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        scriptTag.type = "text/javascript";
        document.head.appendChild(scriptTag);

        function Calculate1() {
            //Calculate attendance rate 計算出席率
            /*
            "在學員動態管理>>教務管理>>不預告實地抽訪紀錄表，由系統自動帶出應到人數及計算出席率
            公式：
            應到人數=參訓人數-退訓人數
            AuthCount=Hid_StudCount-RejectCount
            出席率=(實到人數+請假人數)/應到人數 "
            AtteRate=(TurthCount+TurnoutCount) /AuthCount            
            */
            var Hid_StudCount = $("#Hid_StudCount");//參訓人數
            var AuthCount = $("#AuthCount");//應到人數
            var RejectCount = $("#RejectCount");//退訓人數
            if (Hid_StudCount.val() != "" && RejectCount.val() != "") {
                var x1 = parseInt(Hid_StudCount.val(), 10) - parseInt(RejectCount.val(), 10);
                //AuthCount.val(x1);
            }
            //AuthCount.prop('readonly', true);//應到人數
            //RejectCount.prop('readonly', true);//退訓人數
            //var uumsg = "Calculate1:" + "\nAuthCount" + AuthCount.val() + "\nHid_StudCount" + Hid_StudCount.val() + "\nRejectCount" + RejectCount.val();
            //alert(uumsg);
        }

        function Calculate2() {
            /*
            "在學員動態管理>>教務管理>>不預告實地抽訪紀錄表，由系統自動帶出應到人數及計算出席率
            公式：
            應到人數=參訓人數-退訓人數
            AuthCount=Hid_StudCount-RejectCount
            出席率=(實到人數+請假人數)/應到人數 "
            AtteRate=(TurthCount+TurnoutCount) /AuthCount            
            */
            //debugger;
            var AtteRate = $("#AtteRate");//出席率
            var TurthCount = $("#TurthCount");//實到人數
            var TurnoutCount = $("#TurnoutCount");//請假人數
            var AuthCount = $("#AuthCount");//應到人數
            if (TurthCount.val() != "" && TurnoutCount.val() != "" && AuthCount.val() != "") {
                var x2 = parseInt((parseFloat(TurthCount.val()) + parseFloat(TurnoutCount.val())) / parseFloat(AuthCount.val()) * 100, 10);
                AtteRate.val(x2);
            }
            //AtteRate.prop('readonly', true);//出席率
            //AuthCount.prop('readonly', true);//應到人數
            //var uumsg = "Calculate2:" + "\nAtteRate" + AtteRate.val() + "\nTurthCount" + TurthCount.val() + "\nTurnoutCount" + TurnoutCount.val() + "\nAuthCount" + AuthCount.val();
            //alert(uumsg);
        }

        function showTR() {
            if (document.form1.LItem2.checked) {
                document.getElementById('LItem_TR').style.display = ''; //inline
                document.getElementById('LItem_TR2').style.display = ''; //inline
            }
            else if (document.form1.LItem1.checked) {
                document.form1.LItem2_1.checked = false;
                document.form1.LItem2_2.checked = false;
                document.getElementById('LItem2_1_Date').value = '';
                document.getElementById('LItem2_2_Note').value = '';
                document.getElementById('LItem_TR').style.display = 'none';
                document.getElementById('LItem_TR2').style.display = 'none';
                //document.getElementByID('LItem2_1').checked = false;
                //document.getElementByID('LItem2_2').checked = false;
            }

            //fix 動態變動顯示內容, 會造成顯示內容超出 iframe 顯示區域被遮掉的情況 
            //if (parent) parent.setMainFrameHeight();
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度
            if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
        }

        function chkdata() {
            var msg = '';
            if (document.form1.RIDValue.value == '') msg += '請選擇機構\n';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇職類\n';
            if (document.form1.ApplyDate.value == '') msg += '請輸入訪查日期\n';
            if (document.form1.OLessonTeah1.value == '') {
                msg += '請輸入當日教師\n';
            }
            else {
                if (document.form1.OLessonTeah1Value.value == '') msg += '請重新選擇當日教師\n';
            }
            if (document.form1.OLessonTeah2.value != '') {
                if (document.form1.OLessonTeah2Value.value == '') msg += '請重新選擇當日助教\n';
            }
            if (document.form1.CourseName.value == '') msg += '請輸入當日課程名稱\n';
            if (bl_rocYear == "Y") {
                if (document.form1.ApplyDate.value != '' && !checkRocDate(document.form1.ApplyDate.value)) msg += '訪查日期的時間格式不正確\n';
            }
            else {
                if (document.form1.ApplyDate.value != '' && !checkDate(document.form1.ApplyDate.value)) msg += '訪查日期的時間格式不正確\n';
            }
            if (document.form1.Applytime_HH.value != '') {
                if (!isUnsignedInt(document.form1.Applytime_HH.value)) msg += '訪查時間的小時格式不正確\n';
                if (document.form1.Applytime_HH.value > 23) msg += '訪查時間的小時格式超過範圍\n';
            }
            else {
                msg += '訪查時間的小時格式必須填寫\n';
            }
            if (document.form1.Applytime_MM.value != '') {
                if (!isUnsignedInt(document.form1.Applytime_MM.value)) msg += '訪查時間的分鐘格式不正確\n';
                if (document.form1.Applytime_MM.value > 59) msg += '訪查時間的分鐘格式超過範圍\n';
            }
            else {
                msg += '訪查時間的分鐘格式必須填寫\n';
            }
            if (document.form1.AuthCount.value != '' && !isUnsignedInt(document.form1.AuthCount.value)) msg += '應到人數必須為數字\n';
            if (document.form1.TurthCount.value != '' && !isUnsignedInt(document.form1.TurthCount.value)) msg += '實到人數必須為數字\n';
            if (document.form1.TruancyCount.value != '' && !isUnsignedInt(document.form1.TruancyCount.value)) msg += '未到人數必須為數字\n';
            if (document.form1.TurnoutCount.value != '' && !isUnsignedInt(document.form1.TurnoutCount.value)) msg += '請假人數必須為數字\n';
            if (document.form1.RejectCount.value != '' && !isUnsignedInt(document.form1.RejectCount.value)) msg += '退訓人數必須為數字\n';
            if (document.form1.OtherCount.value != '' && !isUnsignedInt(document.form1.OtherCount.value)) msg += '其他人數必須為數字\n';
            //if (!isChecked(document.form1.Data1)) msg += '請選擇招訓簡章1的選項\n';
            //if (!isChecked(document.form1.Data2)) msg += '請選擇學員出缺席管理1的選項\n';
            //if (!isChecked(document.form1.Data3)) msg += '請選擇教學日誌1的選項\n';
            if (!isChecked(document.form1.DATA81)) msg += '請選擇 一、資料文件查核-1.學員簽到(退)及教學日誌\n';
            if (!isChecked(document.form1.DATA82)) msg += '請選擇 一、資料文件查核-2.訓練課程開班學員名冊\n';
            if (!isChecked(document.form1.DATA83)) msg += '請選擇 一、資料文件查核-3.當日學員請假相關證明\n';
            if (!isChecked(document.form1.DATA84)) msg += '請選擇 一、資料文件查核-4.視需要提供當日「學員領取材料(書籍)單」\n';
            if (!isChecked(document.form1.Item1)) msg += '請回答課程內容是否變動?\n';
            if (!isChecked(document.form1.Item2)) msg += '請回答是否正常上課?\n';
            if (!isChecked(document.form1.Item3)) msg += '請回答訓練場地是否有異動?\n';
            if (!isChecked(document.form1.Item4)) msg += '請回答抽訪學員是否與名單相符?\n';
            if (document.form1.Stud_Name.value == '') msg += '請輸入抽訪學員之姓名1\n';
            if (document.form1.Stud_Name2.value == '') msg += '請輸入抽訪學員之姓名2\n';
            if (!isChecked(document.form1.SItem1)) msg += '請回答抽訪學員是否相符?\n';
            if (!isChecked(document.form1.SItem2)) msg += '請回答實際上課學員人數簽到表是否單相符?\n';
            //if (!isChecked(document.form1.SItem3)) msg += '請回答招訓簡章之內容是否符合作業手冊之相關規定?\n';
            if (document.form1.CurseName.value == '') msg += '請輸入培訓單位人員姓名?\n';
            if (document.form1.VisitorName.value == '') msg += '請輸入訪視人員姓名?\n';
            if (!document.form1.LItem1.checked && !document.form1.LItem2.checked) { msg += '請選取抽訪結果!!\n'; }
            //if (!document.form1.LItem1.checked && !document.form1.LItem2.checked && !document.form1.LItem3.checked) msg += '請選取抽訪結果!!\n';
            // if (document.form1.LItem1.checked && document.form1.LItem2.checked) msg+='請擇一選取抽訪結果!!\n'; 
            if (document.form1.LItem2.checked) {
                if (!document.form1.LItem2_1.checked && !document.form1.LItem2_2.checked) {
                    msg += '請擇一選取抽訪需修正原因!!\n';
                }
                if (document.form1.LItem2_1.checked && document.form1.LItem2_1_Date.value == '') {
                    msg += '請選取限期補正日期!!\n';
                }
                if (document.form1.LItem2_2.checked && document.form1.LItem2_2_Note.value == '') {
                    msg += '請輸入需修正其他原因!!\n';
                }
            }
            if (checkMaxLen(document.getElementById('SItem1_Note').value, 100 * 2)) {
                msg += '【1.抽訪學員之姓名 其他說明】長度不可超過100字元\n';
            }
            if (checkMaxLen(document.getElementById('SItem2_Note').value, 100 * 2)) {
                msg += '【2.實際上課學員人數簽到表 其他說明】長度不可超過100字元\n';
            }
            /*
			if (checkMaxLen(document.getElementById('SItem3_Note').value, 100 * 2)) {
			msg += '【3.招訓簡章之內容是否符合作業手冊之相關規定 其他說明】長度不可超過100字元\n';
			}
			*/
            if (checkMaxLen(document.getElementById('LItem2_2_Note').value, 100 * 2)) {
                msg += '【四、現場處理說明： 其他說明】長度不可超過100字元\n';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function OpenLessonTeah1(opentype) {
            wopen('../../CP/01/ListTeah.aspx?RID=' + document.form1.RIDValue.value
					+ '&Channel=CP_01_006_add8'
					+ '&type=' + opentype
					+ '&fieldname=1'
					+ '&OLessonTeah=OLessonTeah1'
					+ '&OLessonTeahValue=OLessonTeah1Value'
					+ '&ODegreeID=ODegreeID1'
					+ '&ODegreeIDValue=ODegreeIDValue1'
					, 'LessonTeah1', 780, 530, 1);
        }

        function OpenLessonTeah2(opentype) {
            wopen('../../CP/01/ListTeah.aspx?RID=' + document.form1.RIDValue.value
					+ '&Channel=CP_01_006_add8'
					+ '&type=' + opentype
					+ '&fieldname=2'
					+ '&OLessonTeah=OLessonTeah2'
					+ '&OLessonTeahValue=OLessonTeah2Value'
					+ '&ODegreeID=ODegreeID2'
					+ '&ODegreeIDValue=ODegreeIDValue2'
					, 'LessonTeah2', 780, 530, 1);
        }

        function LessonCourID1(opentype, fieldname) {
            if (document.form1.RIDValue.value != '' && document.form1.OCIDValue1.value != '' && document.form1.ApplyDate.value != '') {
                wopen('./ListCourID.aspx?RID=' + document.form1.RIDValue.value + '&OCID=' + document.form1.OCIDValue1.value + '&ApplyDate=' + document.form1.ApplyDate.value + '&type=' + opentype + '&fieldname=', 'LessonCourID1', 780, 530, 1);
            }
            else {
                alert('請先選擇機構、職類/班別與訪查時間');
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;不預告實地抽訪紀錄表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Table1" border="0" cellspacing="1" cellpadding="1" width="100%" class="table_sch">
            <tr>
                <td>

                    <table id="Table3" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol" style="width: 20%">機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="40%"></asp:TextBox>
                                <input id="Button2" value="..." type="button" name="Button2" runat="server" class="button_b_Mini">
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server"><br>
                                <span style="position: absolute; display: none" id="HistoryList2">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol"><font>職類/班別</font> </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input id="Button3" onclick="javascript: window.open('../CP_01_ch.aspx?RID=' + document.form1.RIDValue.value, '', 'width=650,height=650,location=0,status=0,menubar=0,scrollbars=1,resizable=0');" value="..." type="button" name="Button3" runat="server" class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server"><br>
                                <span style="position: absolute; display: none; left: 28%" id="HistoryList">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol"><font>訪查時間</font></td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ApplyDate" runat="server" onfocus="this.blur()" Width="18%" MaxLength="10"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('ApplyDate','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                                <input style="width: 8%" id="Applytime_HH" maxlength="2" runat="server">點：
							    <input style="width: 8%" id="Applytime_MM" maxlength="2" runat="server">分
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="width: 20%">當日教師</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox Style="cursor: pointer" ID="OLessonTeah1" onfocus="this.blur()" Width="50%" MaxLength="8" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                                <input id="OLessonTeah1Value" type="hidden" name="OLessonTeah1Value" runat="server">
                                <input id="teacherbtn" value="..." type="button" name="teacherbtn" runat="server" class="button_b_Mini">
                                <input id="ODegreeID1" type="hidden" name="ODegreeID1" runat="server">
                                <input id="ODegreeIDValue1" type="hidden" name="ODegreeIDValue1" runat="server">
                            </td>
                            <td class="bluecol" style="width: 20%">當日助教 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox Style="cursor: pointer" ID="OLessonTeah2" onfocus="this.blur()" Width="50%" MaxLength="8" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師"></asp:TextBox>
                                <input id="OLessonTeah2Value" type="hidden" name="OLessonTeah2Value" runat="server">
                                <input id="teacherbtn2" value="..." type="button" name="teacherbtn2" runat="server" class="button_b_Mini">
                                <input id="ODegreeID2" type="hidden" name="ODegreeID2" runat="server">
                                <input id="ODegreeIDValue2" type="hidden" name="ODegreeIDValue2" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol"><font>當日課程</font> </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="CourseName" runat="server" MaxLength="100" Columns="55" Width="60%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">學員出缺勤狀況</td>
                            <td colspan="3">
                                <%--參訓人數--%>
                                <table id="Table4" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td class="whitecol">應到人數:<asp:TextBox ID="AuthCount" runat="server" Width="30%"></asp:TextBox>人 </td>
                                        <td class="whitecol">實到人數:<asp:TextBox ID="TurthCount" runat="server" Width="30%"></asp:TextBox>人 </td>
                                        <td class="whitecol">未到人數:<asp:TextBox ID="TruancyCount" runat="server" Width="30%"></asp:TextBox>人 </td>
                                        <td class="whitecol">出席率:<asp:TextBox ID="AtteRate" runat="server" Width="30%" MaxLength="6" onfocus="this.blur()"></asp:TextBox>%</td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">退訓人數:<asp:TextBox ID="RejectCount" runat="server" Width="30%"></asp:TextBox>人 </td>
                                        <td class="whitecol">請假人數:<asp:TextBox ID="TurnoutCount" runat="server" Width="30%"></asp:TextBox>人 </td>
                                        <td class="whitecol" colspan="2">其他人數:<asp:TextBox ID="OtherCount" runat="server" Width="15%"></asp:TextBox>人 </td>
                                    </tr>
                                    <tr>
                                        <td colspan="4" class="whitecol">點名未到課學員，另以電話抽訪 </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <div>以下由訪視人員填寫</div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">【一、資料文件查核】</td>
                            <td colspan="3">
                                <table id="Table6" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td width="33%" class="whitecol">1.學員簽到(退)及教學日誌</td>
                                        <td width="33%" class="whitecol">
                                            <asp:RadioButtonList ID="DATA81" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">齊</asp:ListItem>
                                                <asp:ListItem Value="2">缺</asp:ListItem>
                                                <asp:ListItem Value="5">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td width="33%" class="whitecol">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">2.訓練課程開班學員名冊</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="DATA82" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">齊</asp:ListItem>
                                                <asp:ListItem Value="2">缺</asp:ListItem>
                                                <asp:ListItem Value="5">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td width="33%" class="whitecol">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">3.當日學員請假相關證明</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="DATA83" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">齊</asp:ListItem>
                                                <asp:ListItem Value="2">缺</asp:ListItem>
                                                <asp:ListItem Value="5">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td width="33%" class="whitecol">&nbsp;</td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">4.視需要提供當日「學員領取材料(書籍)單」</td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="DATA84" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">齊</asp:ListItem>
                                                <asp:ListItem Value="2">缺</asp:ListItem>
                                                <asp:ListItem Value="5">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td width="33%" class="whitecol">&nbsp;</td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">【二、課程執行情形】</td>
                            <td colspan="3">
                                <table id="Table6b" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                    <tr>
                                        <td width="33%" class="whitecol">1.課程內容是否變動？ </td>
                                        <td width="33%" class="whitecol">
                                            <asp:RadioButtonList ID="Item1" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td rowspan="4" width="33%" class="whitecol"></td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">2.是否正常上課？ </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item2" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">3.訓練場地是否有異動？ </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="Item3" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">4.抽訪學員是否與名單相符？ </td>
                                        <td class="whitecol"><font face="新細明體">
                                            <asp:RadioButtonList ID="Item4" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">是</asp:ListItem>
                                                <asp:ListItem Value="2">否</asp:ListItem>
                                                <asp:ListItem Value="3">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </font></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">三、現場訪查實況：</td>
                            <td colspan="3">
                                <table id="Table6c" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">

                                    <tr>
                                        <td width="33%" class="whitecol">1.抽訪學員之姓名1<font color="#ff0066"><strong>*</strong></font>
                                            <asp:TextBox ID="Stud_Name" runat="server" Width="50%" MaxLength="10"></asp:TextBox><br />
                                            &nbsp; &nbsp;抽訪學員之姓名2<font color="#ff0066"><strong>*</strong></font>
                                            <asp:TextBox ID="Stud_Name2" runat="server" Width="50%" MaxLength="10"></asp:TextBox>
                                        </td>
                                        <td width="33%" class="whitecol">
                                            <asp:RadioButtonList ID="SItem1" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                <asp:ListItem Value="1">相符</asp:ListItem>
                                                <asp:ListItem Value="2">不相符</asp:ListItem>
                                                <asp:ListItem Value="3">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td width="33%" class="whitecol">其他說明<asp:TextBox ID="SItem1_Note" runat="server" MaxLength="100" Height="40px" TextMode="MultiLine" Width="70%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="whitecol">2.實際上課學員人數簽到表 </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="SItem2" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                <asp:ListItem Value="1">相符</asp:ListItem>
                                                <asp:ListItem Value="2">不相符</asp:ListItem>
                                                <asp:ListItem Value="3">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">其他說明<asp:TextBox ID="SItem2_Note" runat="server" MaxLength="100" Height="40px" TextMode="MultiLine" Width="70%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <%--<tr>
                                        <td class="whitecol">
                                            3.招訓簡章之內容是否符合作業手冊之相關規定
                                        </td>
                                        <td class="whitecol">
                                            <asp:RadioButtonList ID="SItem3" runat="server" Width="176px" CssClass="font" RepeatDirection="Horizontal">
                                                <asp:ListItem Value="1">相符</asp:ListItem>
                                                <asp:ListItem Value="2">不相符</asp:ListItem>
                                                <asp:ListItem Value="3">其他</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="SItem3_Note" runat="server" Width="100%" MaxLength="100" Height="37px" TextMode="MultiLine"></asp:TextBox>
                                        </td>
                                    </tr>--%>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">四、現場處理說明：</td>
                            <td colspan="3">
                                <table id="Table6d" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                                    <%--<tr>
                                        <td colspan="3" class="bluecol">
                                            <p align="center"><font><font color="#ff0066"><strong><font color="#ffff66"><strong>* </strong></font></strong></font>四、現場處理說明：</font>&nbsp;</p>
                                        </td>
                                    </tr>--%>
                                    <tr>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButton ID="LItem1" onclick="showTR();" runat="server" Text="1.不預告抽訪結果尚屬正常" GroupName="LItem"></asp:RadioButton>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="3" class="whitecol">
                                            <asp:RadioButton ID="LItem2" onclick="showTR();" runat="server" Text="2.不預告抽訪結果需修正如下：" GroupName="LItem"></asp:RadioButton>
                                        </td>
                                    </tr>
                                    <tr id="LItem_TR" runat="server">
                                        <td colspan="3" class="whitecol">&nbsp;&nbsp;&nbsp;(1)<asp:CheckBox ID="LItem2_1" runat="server"></asp:CheckBox>學員資料有誤或填寫錯誤，(已影印學員名冊存查)，需請訓練單位於
                                            <asp:TextBox ID="LItem2_1_Date" runat="server" onfocus="this.blur()" Width="18%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('LItem2_1_Date','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">前補正相關資料，並傳真至轄區 分署 核備。
                                        </td>
                                    </tr>
                                    <tr id="LItem_TR2" runat="server">
                                        <td colspan="1" class="whitecol" nowrap="nowrap">&nbsp;&nbsp;&nbsp;(2)<asp:CheckBox ID="LItem2_2" runat="server"></asp:CheckBox>其他(複選)： </td>
                                        <td colspan="2" class="whitecol">
                                            <asp:CheckBoxList ID="cblLItem2_2b" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="5">
                                                <asp:ListItem Value="01">出席率不佳</asp:ListItem>
                                                <asp:ListItem Value="02">簽到退未落實</asp:ListItem>
                                                <asp:ListItem Value="03">師資不符</asp:ListItem>
                                                <asp:ListItem Value="06">助教不符</asp:ListItem>
                                                <asp:ListItem Value="04">課程內容不符</asp:ListItem>
                                                <asp:ListItem Value="05">上課地點不符</asp:ListItem>
                                                <asp:ListItem Value="99">其他：</asp:ListItem>
                                            </asp:CheckBoxList>
                                            <asp:TextBox ID="LItem2_2_Note" runat="server" MaxLength="100" Width="98%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td colspan="3" class="whitecol">&nbsp;&nbsp;&nbsp;&nbsp;3.其他補充說明<asp:TextBox ID="LItem2_3_Note" runat="server" Width="84%" MaxLength="300" Height="40px" TextMode="MultiLine"></asp:TextBox></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="15%">五、重要工作事項未依核定課程施訓： </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBoxList ID="cblSITEM51" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                <asp:TextBox ID="SITEM51_NOTE" runat="server" MaxLength="300" Width="60%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="15%">六、課程異常狀況： </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBoxList ID="cblSITEM61" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                <asp:TextBox ID="SITEM61_NOTE" runat="server" MaxLength="300" Width="60%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="15%">七、其他未依核定課程施訓： </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBoxList ID="cblSITEM71" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                <asp:TextBox ID="SITEM71_NOTE" runat="server" MaxLength="300" Width="60%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="15%">八、其他重大異常狀況： </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBoxList ID="cblSITEM81" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                                <asp:TextBox ID="SITEM81_NOTE" runat="server" MaxLength="300" Width="60%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="15%">未實際訪視 </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBox ID="chkB_NOINC5" runat="server" Text="不列入訪視計次" /></td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol">訪視人員單位 </td>
                            <td colspan="3" class="whitecol">
                                <asp:TextBox ID="OrgName" runat="server" onfocus="this.blur()" Width="50%" MaxLength="50"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td width="15%" class="bluecol">培訓單位人員姓名 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="CurseName" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                            <td width="15%" class="bluecol">訪視人員姓名 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="VisitorName" runat="server" MaxLength="50" Width="50%"></asp:TextBox></td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <%--&nbsp;<input id="Button4" value="回查詢頁面" type="button" name="Button4" runat="server" class="button_b_M">--%>
                        <input id="Button5" value="回查詢頁面" type="button" name="Button4" runat="server" class="asp_button_M">
                    </div>
                </td>
            </tr>
        </table>
        <input id="OrgID" type="hidden" runat="server" />
        <asp:HiddenField ID="Hid_StudCount" runat="server" />
        <asp:HiddenField ID="Hid_State1" runat="server" />
        <%--<asp:HiddenField ID="Hid_VerDate" runat="server" />--%>
    </form>
</body>
</html>
