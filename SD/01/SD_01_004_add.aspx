<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_004_add.aspx.vb" Inherits="WDAIIP.SD_01_004_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>e網報名審核</title>
    <meta name="generator" content="microsoft visual studio .net 7.1" />
    <meta name="code_language" content="visual basic .net 7.1" />
    <meta name="vs_defaultclientscript" content="javascript" />
    <meta name="vs_targetschema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
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
        function Change() {
            var ddlBudID = document.getElementById('ddlBudID');
            var ddlSupplyID = document.getElementById('ddlSupplyID');
            if (ddlBudID && ddlSupplyID) {
                if (ddlBudID.value == '97') { ddlSupplyID.value = '2'; }
                else if (ddlBudID.value == '99') { ddlSupplyID.value = '9'; }
                else { ddlSupplyID.value = '請選擇'; }
            }
        }

        function CheckData() {
            //'審核失敗 CheckData
            var msg = "";
            var signUpMemo = document.getElementById('signUpMemo');
            if (signUpMemo && signUpMemo.value == '') {
                msg += '請輸入失敗原因\n';
            }
            else {
                if (checkMaxLen(signUpMemo.value, 150 * 2)) {
                    msg += '【失敗原因】長度不可超過150字元\n';
                }
            }
            if (msg != "") {
                alert(msg);
                return false;
            }
        }

        function CheckData1() {
            //debugger;'審核成功 CheckData1
            var msg = "";
            var signUpMemo = document.getElementById('signUpMemo');
            var IDNO = document.getElementById('IDNO');
            var IDNOValue = document.getElementById('IDNOValue');
            if (signUpMemo && signUpMemo.value != '') {
                if (checkMaxLen(signUpMemo.value, 150 * 2)) {
                    msg += '【失敗原因】長度不可超過150字元\n';
                }
            }
            if (IDNO && IDNOValue) {
                if (IDNO.innerText == "" || IDNOValue.value == "") {
                    msg += "身分證字號不可為空值，請連絡 e網系統管理者";
                    //alert('身分證字號不可為空值，請連絡 e網系統管理者'); //return false;
                }
            }
            if (msg != "") {
                alert(msg);
                return false;
            }
            var rst2 = true; //正常再次檢核。
            var Hid_MSG1 = document.getElementById('Hid_MSG1');
            var Hid_MSG2 = document.getElementById('Hid_MSG2');
            if (Hid_MSG1 && rst2 && Hid_MSG1.value != '') {
                rst2 = confirm(Hid_MSG1.value);
            }
            if (Hid_MSG2 && rst2 && Hid_MSG2.value != '') {
                rst2 = confirm(Hid_MSG2.value);
            }
            return rst2;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%--,<table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">,<tr>,<td>,<asp:Label ID="TitleLab1" runat="server"></asp:Label>
           ,<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;e網報名審核</asp:Label>,</td>,</tr>,</table>,--%>
        <table id="Table2" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td align="center">
                    <table id="Table1" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width: 20%">姓名 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:Label ID="Name" runat="server"></asp:Label></td>
                            <td class="bluecol" style="width: 20%">出生日期 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:Label ID="Birthday" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">身分別 </td>
                            <td class="whitecol">
                                <asp:Label ID="PassPortNO" runat="server"></asp:Label></td>
                            <td class="bluecol">身分證號碼 </td>
                            <td class="whitecol">
                                <asp:Label ID="IDNO" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">性別 </td>
                            <td class="whitecol">
                                <asp:Label ID="Sex" runat="server"></asp:Label></td>
                            <td class="bluecol">最高學歷 </td>
                            <td class="whitecol">
                                <asp:Label ID="DegreeID" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">婚姻狀況 </td>
                            <td class="whitecol">
                                <asp:Label ID="MaritalStatus" runat="server"></asp:Label></td>
                            <td class="bluecol">畢業狀況 </td>
                            <td class="whitecol">
                                <asp:Label ID="GradID" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">學校名稱 </td>
                            <td class="whitecol">
                                <asp:Label ID="School" runat="server"></asp:Label></td>
                            <td class="bluecol">科系 </td>
                            <td class="whitecol">
                                <asp:Label ID="Department" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">兵役 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="MilitaryID" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">通訊地址 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Address" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">戶籍地址 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="LabHouseholdAddress" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">聯絡電話(日) </td>
                            <td class="whitecol">
                                <asp:Label ID="Phone1" runat="server"></asp:Label></td>
                            <td class="bluecol">聯絡電話(夜) </td>
                            <td class="whitecol">
                                <asp:Label ID="Phone2" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">電子信箱 </td>
                            <td class="whitecol">
                                <asp:Label ID="Email" runat="server"></asp:Label></td>
                            <td class="bluecol">行動電話 </td>
                            <td class="whitecol">
                                <asp:Label ID="CellPhone" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">主要參訓身分別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Lab1311" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">投保單位名稱 </td>
                            <td class="whitecol">
                                <asp:Label ID="Lab1301" runat="server"></asp:Label></td>
                            <td class="bluecol">投保單位保險證號 </td>
                            <td class="whitecol">
                                <asp:Label ID="Lab1302" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">投保單位類別 </td>
                            <td class="whitecol">
                                <asp:Label ID="Lab1303" runat="server"></asp:Label></td>
                            <td class="bluecol">投保單位電話 </td>
                            <td class="whitecol">
                                <asp:Label ID="Lab1304" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">投保單位地址 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Lab1305" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">目前公司名稱 </td>
                            <td class="whitecol">
                                <asp:Label ID="Lab1307" runat="server"></asp:Label></td>
                            <td class="bluecol">統一編號 </td>
                            <td class="whitecol">
                                <asp:Label ID="Lab1308" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">目前任職部門 </td>
                            <td class="whitecol">
                                <asp:Label ID="Lab1309" runat="server"></asp:Label></td>
                            <td class="bluecol">職稱 </td>
                            <td class="whitecol">
                                <asp:Label ID="Lab1310" runat="server"></asp:Label></td>
                        </tr>
                        <tr id="TRIdentityID" runat="server">
                            <td class="bluecol">參訓身分別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="IdentityID" runat="server"></asp:Label></td>
                        </tr>
                        <tr id="TRHandTypeID" runat="server">
                            <td class="bluecol">障礙類別 </td>
                            <td class="whitecol">
                                <asp:Label ID="labHandTypeID" runat="server" Width="136px"></asp:Label></td>
                            <td class="bluecol">障礙等級 </td>
                            <td class="whitecol">
                                <asp:Label ID="labHandLevelID" runat="server" Width="120px"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">報名日期 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="RelEnterDate" runat="server"></asp:Label></td>
                        </tr>
                        <%--<tr><td class="bluecol">報名志願 </td><td class="whitecol" colspan="3">
                            <table id="Table4" class="font" border="0" cellspacing="1" cellpadding="1">
                        <tr><td>第一志願： </td><td><asp:Label ID="OCID1" runat="server"></asp:Label></td></tr>
                        <tr><td>第二志願： </td><td><asp:Label ID="OCID2" runat="server"></asp:Label></td></tr>
                        <tr><td>第三志願： </td><td><asp:Label ID="OCID3" runat="server"></asp:Label></td></tr></table></td></tr>--%>
                        <tr>
                            <td class="bluecol">報名班級 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="OCID1" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="Table11" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width: 20%">報名班級 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:Label ID="Label1" runat="server"></asp:Label></td>
                            <td class="bluecol" style="width: 20%">報名日期 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:Label ID="Label2" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">姓名 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label3" runat="server"></asp:Label></td>
                            <td class="bluecol">出生日期 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label4" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">身分別 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label5" runat="server"></asp:Label></td>
                            <td class="bluecol">身分證號碼 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label6" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">性別 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label7" runat="server"></asp:Label></td>
                            <td class="bluecol">最高學歷 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label9" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">通訊地址 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Label14" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">戶籍地址 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Label15" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">聯絡電話(日) </td>
                            <td class="whitecol">
                                <asp:Label ID="Label16" runat="server"></asp:Label></td>
                            <td class="bluecol">聯絡電話(夜) </td>
                            <td class="whitecol">
                                <asp:Label ID="Label17" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">電子信箱 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label18" runat="server"></asp:Label></td>
                            <td class="bluecol">行動電話 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label19" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">主要參訓身分別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Label20" runat="server"></asp:Label></td>
                        </tr>
                        <%--<tr><td class="bluecol">受訓前薪資 </td><td class="whitecol" colspan="3"><asp:Label ID="Label30" runat="server"></asp:Label></td></tr>
                            <tr><td class="bluecol">郵政/銀行帳號 </td><td class="whitecol" colspan="3"><asp:Label ID="Label33" runat="server"></asp:Label></td></tr>--%>
                    </table>
                    <%--<table id="Table22" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                    <tr><td class="bluecol" style="width: 20%">局號 </td><td class="whitecol" style="width: 30%"><asp:Label ID="Label35" runat="server"></asp:Label></td><td class="bluecol" style="width: 20%">帳號 </td><td class="whitecol" style="width: 30%"><asp:Label ID="Label36" runat="server"></asp:Label></td></tr>
                    </table>
                    <table id="Table23" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                    <tr><td class="bluecol" style="width: 20%">總行名稱 </td><td class="whitecol" style="width: 30%"><asp:Label ID="Label37" runat="server"></asp:Label></td><td class="bluecol" style="width: 20%">總行代號 </td>
                        <td class="whitecol" style="width: 30%"><asp:Label ID="Label38" runat="server"></asp:Label></td></tr><tr><td class="bluecol">分行名稱 </td><td class="whitecol"><asp:Label ID="Label61" runat="server">
                        </asp:Label></td><td class="bluecol">分行代號 </td><td class="whitecol"><asp:Label ID="Label62" runat="server"></asp:Label></td></tr><tr><td class="bluecol">帳號 </td><td class="whitecol" colspan="3">
                        <asp:Label ID="Label34" runat="server"></asp:Label></td></tr>
                    </table>--%>

                    <table id="Table12" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width: 20%">投保單位名稱 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:Label ID="Label59" runat="server"></asp:Label></td>
                            <td class="bluecol" style="width: 20%">投保單位保險證號 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:Label ID="ACTNO60" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">投保單位類別 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label63" runat="server"></asp:Label></td>
                            <td class="bluecol">投保單位電話 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label64" runat="server"></asp:Label></td>
                        </tr>
                        <%--<tr><td class="bluecol">投保單位<br />統一編號</td><td class="whitecol" colspan="3"><asp:Label id="actcomIDNO" runat="server"></asp:Label></td></tr>--%>
                        <tr>
                            <td class="bluecol">投保單位地址 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Label65" runat="server"></asp:Label></td>
                        </tr>
                        <%--<tr><td class="bluecol">第一次投保日</td><td class="whitecol" colspan="3"><asp:Label id="Label39" runat="server" width="136px"></asp:Label></td></tr>--%>
                        <tr>
                            <td class="bluecol">目前公司名稱 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label40" runat="server"></asp:Label></td>
                            <td class="bluecol">統一編號 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label41" runat="server"></asp:Label></td>
                        </tr>
                        <%--<tr><td class="bluecol">公司電話</td><td class="whitecol"><asp:Label id="Label42" runat="server" width="144px"></asp:Label></td><td class="bluecol">公司傳真</td><td class="whitecol">
                            <asp:Label id="Label43" runat="server" width="152px"></asp:Label></td></tr><tr><td class="bluecol">公司地址</td><td class="whitecol" colspan="3">
                            <asp:Label id="Label44" runat="server" width="424px"></asp:Label></td></tr>--%>
                        <tr>
                            <td class="bluecol">目前任職部門 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label45" runat="server"></asp:Label></td>
                            <td class="bluecol">職稱 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label46" runat="server"></asp:Label></td>
                        </tr>
                        <%--<tr><td class="bluecol">個人到任目前任職公司起日</td><td class="whitecol"><asp:Label id="Label47" runat="server" width="144px"></asp:Label></td><td class="bluecol">個人到任目前職務起日</td>
                       <td class="whitecol"><asp:Label id="Label48" runat="server" width="152px"></asp:Label></td></tr><tr><td class="bluecol">最近升遷日期</td><td class="whitecol">
                       <asp:Label id="Label49" runat="server" width="144px"></asp:Label></td><td class="bluecol">是否由公司推薦參訓</td><td class="whitecol">
                       <asp:Label id="Label50" runat="server" width="152px"></asp:Label></td></tr>--%>
                        <tr>
                            <td class="bluecol">是否由公司推薦參訓 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Label50" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">參訓動機 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Label51" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">服務單位行業別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="Label58" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓後動向 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label52" runat="server"></asp:Label></td>
                            <td class="bluecol">服務單位是否<br />
                                屬於中小企業 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label53" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">個人工作年資 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label54" runat="server"></asp:Label></td>
                            <td class="bluecol">在這家公司的年資 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label55" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">在這職位的年資 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label56" runat="server"></asp:Label></td>
                            <td class="bluecol">最近升遷<br />
                                離本職幾年 </td>
                            <td class="whitecol">
                                <asp:Label ID="Label57" runat="server"></asp:Label></td>
                        </tr>
                        <%--<tr><td class="bluecol">預算別 </td><td class="whitecol"><asp:DropDownList ID="ddlBudID" runat="server"><asp:ListItem Value="01">公務</asp:ListItem><asp:ListItem Value="02">就安</asp:ListItem>
                          <asp:ListItem Value="03">就保</asp:ListItem><asp:ListItem Value="97">協助</asp:ListItem><asp:ListItem Value="99">不補助</asp:ListItem>
                          <asp:ListItem Value="請選擇">請選擇</asp:ListItem></asp:DropDownList></td><td class="bluecol">補助比例 </td><td class="whitecol"><asp:DropDownList ID="ddlSupplyID" runat="server">
                          <asp:ListItem Value="1">一般80%</asp:ListItem><asp:ListItem Value="2">特定100%</asp:ListItem><asp:ListItem Value="9">0%</asp:ListItem><asp:ListItem>請選擇</asp:ListItem></asp:DropDownList></td></tr>--%>
                        <tr id="trTPlanID28DBL2" runat="server">
                            <td class="bluecol">補助費用 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labMoneyShow1" runat="server"></asp:Label>
                                <br />
                                <asp:Label ID="labOver6w" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr id="trTPlanID28DBL3" runat="server">
                            <td class="bluecol">補助費用說明 </td>
                            <td class="whitecol" colspan="3">
                                <asp:Label ID="labMsg2" runat="server">*預估補助費用是以該課程費用80%作為估算</asp:Label></td>
                        </tr>
                    </table>
                    <%-- <table id="Table13" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server"></table>--%>
                    <table id="TablePWTYPE" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width: 20%">受訓前任職狀況 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:Label ID="PriorWorkType1" runat="server"></asp:Label></td>
                            <td class="bluecol" style="width: 20%">最後一次任職<br />
                                單位名稱 </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:Label ID="PriorWorkOrg1" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">最後投保單位<br />
                                起迄日 </td>
                            <td class="whitecol">
                                <asp:Label ID="OfficeDate" runat="server"></asp:Label></td>
                            <td class="bluecol">最後投保單位<br />
                                保險證號 </td>
                            <td class="whitecol">
                                <asp:Label ID="ActNo" runat="server"></asp:Label></td>
                        </tr>
                    </table>
                    <table id="DataGridTable" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td align="center" class="table_title">已申請失業給付資料</td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" Height="6px" CssClass="font" AutoGenerateColumns="false" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="apply_date" HeaderText="申請給付起日" DataFormatString="{0:d}">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="apply_money" HeaderText="金額">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="station_Name" HeaderText="就業服務站">
                                            <HeaderStyle HorizontalAlign="Center" Width="70%"></HeaderStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable2" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td align="center" class="table_title">已申請職訓生活津貼</td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="orgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="30%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="classcName" HeaderText="參訓課程">
                                            <HeaderStyle Width="30%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="stdate" HeaderText="開訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ftdate" HeaderText="結訓日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="trainingmoney" HeaderText="申請補助金額">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="paymoney" HeaderText="實領核發金額">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table id="Table10" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" width="20%">失敗原因 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="signUpMemo" runat="server" MaxLength="150" TextMode="multiline" Columns="40" Rows="4"></asp:TextBox>
                                <asp:Label ID="lab_LastModifyDate" runat="server" Text=""></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="tbTPlanID28DBL1" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td align="center" class="table_title">產投方案的報名或參訓時段重疊查詢</td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="msgbb" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="DataGrid2bb" runat="server" Width="100%" AllowSorting="True" PageSize="20" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                            <HeaderTemplate>序號</HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="Labsignno" runat="server"></asp:Label>
                                                <asp:Label ID="Labdouble" runat="server" ForeColor="Red">重</asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="DistName" HeaderText="轄區">
                                            <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Years" HeaderText="年度">
                                            <HeaderStyle Width="12%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PlanName" HeaderText="訓練計畫">
                                            <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練單位">
                                            <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="CLASSCNAME" HeaderText="課程名稱"><HeaderStyle HorizontalAlign="Center" Width="70px"></HeaderStyle></asp:BoundColumn>--%>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                            <HeaderTemplate>課程名稱</HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labCLASSCNAME" runat="server"></asp:Label><br />
                                                <asp:HyperLink ID="HrLk1" runat="server" CssClass="newlink">課程連結</asp:HyperLink>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="TRound" HeaderText="訓練期間">
                                            <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                            <HeaderTemplate>(重疊)日期-上課時間</HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Literal ID="Literal1" runat="server"></asp:Literal>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                                            <HeaderTemplate>訓練狀態</HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="Labstudstatus" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table id="Tablehistory3" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td align="center" class="table_title">近兩年參訓資料查詢</td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="DataGrid3" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" PageSize="5" AllowSorting="true" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <%--,<headerstyle forecolor="white" backcolor="#2aafc0"></headerstyle>,<asp:boundcolumn datafield="distName" sortexpression="distName" headertext="轄區&lt;br&gt;中心">
                                        ,<headerstyle HorizontalAlign="Center" width="10%"></headerstyle>,<itemstyle HorizontalAlign="Center"></itemstyle>,</asp:boundcolumn>
                                        ,<asp:boundcolumn datafield="years" headertext="年度">,<headerstyle width="6%"></headerstyle>,</asp:boundcolumn>
                                        ,<asp:boundcolumn datafield="orgName" sortexpression="orgName" headertext="訓練機構">,<headerstyle HorizontalAlign="Center" width="20%"></headerstyle>
                                        ,</asp:boundcolumn>,--%>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="planName" HeaderText="訓練計畫">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="orgName" HeaderText="訓練機構">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="tmid" HeaderText="訓練職類">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="cjob_Name" HeaderText="通俗職類"><HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle></asp:BoundColumn>--%>
                                        <asp:TemplateColumn HeaderText="通俗職類">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="cjob_Name" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="className" HeaderText="班別名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="受訓&lt;br&gt;時數">
                                            <HeaderStyle Width="5%" />
                                            <ItemTemplate>
                                                <asp:Label ID="thours" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="受訓期間">
                                            <HeaderStyle Width="8%" />
                                            <ItemTemplate>
                                                <asp:Label ID="tround" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="skillName" HeaderText="技能檢定">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="weeks" HeaderText="上課時間">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="訓練&lt;br&gt;狀態">
                                            <HeaderStyle Width="5%" />
                                            <ItemTemplate>
                                                <asp:Label ID="tflag" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <%--<asp:BoundColumn DataField="JobStatus" HeaderText="就業狀況"><HeaderStyle Width="5%" /></asp:BoundColumn>--%>
                                    </Columns>
                                    <PagerStyle Visible="false"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="DataGrid34" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" PageSize="5" AllowSorting="true" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="VSSORT" HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PLANNAME" HeaderText="訓練計畫">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構">
                                            <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TMID" HeaderText="訓練職類">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CJOB_NAME" HeaderText="通俗職類">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassName" HeaderText="班別名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="THours" HeaderText="受訓時數">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TRound" HeaderText="受訓期間">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="WEEKS" HeaderText="上課時間">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="TFlag" HeaderText="訓練&lt;br&gt;狀態">
                                            <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="false"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable4" class="Table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td align="center" class="table_title">已報名同一甄試日期</td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <asp:DataGrid ID="DataGrid4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="false" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="planName" HeaderText="訓練計畫">
                                            <HeaderStyle Width="15%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="orgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="15%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="classcName" HeaderText="報名參訓課程">
                                            <HeaderStyle Width="20%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="examdate" HeaderText="甄試日期">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="stdate" HeaderText="開訓日期">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ftdate" HeaderText="結訓日期">
                                            <HeaderStyle Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="signUpStatus2" HeaderText="備註">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="ftdate" HeaderText="結訓日期" DataFormatString="{0:d}"><HeaderStyle Width="80px"></HeaderStyle></asp:BoundColumn>--%>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <div class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="審核成功" CssClass="asp_Button_M"></asp:Button>
                        <asp:Button ID="Button2" runat="server" Text="審核失敗" CssClass="asp_Button_M"></asp:Button>
                        <asp:Button ID="Button3" runat="server" Text="回上一頁" CssClass="asp_Button_M"></asp:Button>
                        <asp:Button ID="Button4" runat="server" Text="回上一頁" Visible="false" CssClass="asp_Button_M"></asp:Button>
                    </div>
                </td>
            </tr>
        </table>
        <input id="IDNOValue" type="hidden" name="IDNOValue" runat="server" />
        <input id="STDateValue" type="hidden" name="STDateValue" runat="server" />
        <input id="BFDateValue" type="hidden" name="BFDateValue" runat="server" />
        <input id="heSerNum" type="hidden" name="heSerNum" runat="server" />
        <input id="heSETID" type="hidden" name="heSETID" runat="server" />
        <asp:HiddenField ID="HidIdentityID" runat="server" />
        <asp:HiddenField ID="HiderrFlag" runat="server" />
        <asp:HiddenField ID="Hid_eSerNum" runat="server" />
        <asp:HiddenField ID="Hid_MSG1" runat="server" />
        <asp:HiddenField ID="Hid_MSG2" runat="server" />
        <asp:HiddenField ID="Hid_MSGADIDN" runat="server" />
        <asp:HiddenField ID="Hid_ACTNObli" runat="server" />
        <asp:HiddenField ID="Hid_PreUseLimited18a" runat="server" />
        <input id="Hid_show_actno_budid" type="hidden" runat="server" />
    </form>
</body>
</html>
