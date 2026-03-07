<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_06_001_chk.aspx.vb" Inherits="WDAIIP.TC_06_001_chk" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>計畫變更審核</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        //檢核送出的資訊
        function fnChkVerify() {
            var strMsg = "";
            //$('#ChkMode').val();
            //$('#ChkMode option:selected').val();
            //var ChkMode = document.getElementById('ChkMode');
            if ($('#ChkMode option:selected').val() == "") {
                strMsg += "請選擇 審核結果!";
            }
            if (strMsg != "") {
                alert(strMsg);
                return false;
            }
            return true;
        }

        //檢核-申請變更函送日期
        function changSTATUS4() {
            var strMsg = "";
            var hid_AltDataID = $("#hid_AltDataID");//9:停辦//15:上課時間//16:其他
            var Hid_ApplyDate = $("#Hid_ApplyDate");
            var SENDDATE4 = $("#SENDDATE4");
            var lab_msg_4 = $("#lab_msg_4");//lab_msg_4 $("#labmsg_4");//lab_msg_4
            //debugger;
            lab_msg_4.html("未檢核!");
            if (SENDDATE4.val() == "") {
                strMsg += "未檢核-請輸入-申請變更函送日期!";
                lab_msg_4.html(strMsg);
                alert(strMsg);
                return false;
            }
            if (Hid_ApplyDate.val() != '' && checkDate(Hid_ApplyDate.val()) && SENDDATE4.val() != '' && checkDate(SENDDATE4.val())) {
                var beginDate = Hid_ApplyDate.val();//new Date();
                var beginDate7 = addDateByDay(beginDate, 7)
                var beginDate14 = addDateByDay(beginDate, 14)
                var endDate = SENDDATE4.val(); //new Date(SENDDATE4.val());
                //OVERWEEK4: 1:1週以內 2:1週以上 3:停辦 9:無逾期
                $('[name="OVERWEEK4"]').removeAttr('checked');
                //STATUS4: 1:依規定辦理 2:逾期(扣分) 3:逾期(不扣分)
                $('[name="STATUS4"]').removeAttr('checked');
                var x3 = getDiffDay(beginDate, endDate);
                if (x3 <= 7) {
                    //lab_msg_4.html("無逾期/依規定辦理" + ",x1:" + x1 + ",x2:" + x2 + ",x3:" + x3);
                    lab_msg_4.html("無逾期-依規定辦理");
                    $("input[name=OVERWEEK4][value=9]").prop('checked', true);
                    $("input[name=STATUS4][value=1]").prop('checked', true);
                }
                if (x3 > 7 && x3 <= 14) {
                    lab_msg_4.html("1周以內-逾期(扣分)");
                    $("input[name=OVERWEEK4][value=1]").prop('checked', true);
                    $("input[name=STATUS4][value=2]").prop('checked', true);
                }
                if (x3 > 14) {
                    lab_msg_4.html("1周以上-逾期(扣分)");
                    $("input[name=OVERWEEK4][value=2]").prop('checked', true);
                    $("input[name=STATUS4][value=2]").prop('checked', true);
                }

                if (hid_AltDataID.val() == "9") {
                    lab_msg_4.html("停辦");
                    $("input[name=OVERWEEK4][value=3]").prop('checked', true);
                }
                if (hid_AltDataID.val() == "15") {
                    lab_msg_4.html("上課時間");
                    //$('[name="OVERWEEK4"]').removeAttr('checked');
                    $("input[name=OVERWEEK4][value=9]").prop('checked', true);
                }
                if (hid_AltDataID.val() == "16") {
                    lab_msg_4.html("其他");
                    //$('[name="OVERWEEK4"]').removeAttr('checked');
                    $("input[name=OVERWEEK4][value=9]").prop('checked', true);
                }
            }
            return false;
        }

        $(function () {
            // Handler for .ready() called.
            $("#SENDDATE4").on("click", function () {
                return changSTATUS4();
                //changSTATUS4();
            });
            $("#btnChkSend4").on("click", function () {
                return changSTATUS4();
                //changSTATUS4();
            });
        });

    </script>
    <%--
/* --
$(document).ready(function () {
    // Handler for .ready() called.
    $("#SENDDATE4").on("click", function () {
        return changSTATUS4();
        //changSTATUS4();
    });
    $("#btnChkSend4").on("click", function () {
        return changSTATUS4();
        //changSTATUS4();
    });
});
*/
    <style type="text/css">
        .bluecol {
            color: Black;
            text-align: right;
            padding: 4px 6px;
            background-color: #f1f9fc;
            border-right: 3px solid #49cbef;
            width: 104px;
        }
    </style>
    --%>
</head>
<body>
    <form id="form1" method="post" runat="server">

        <table id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" style="width: 16%">訓練機構</td>
                            <td class="whitecol" style="width: 34%">
                                <asp:Label ID="OrgName" runat="server"></asp:Label>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" /></td>
                            <td class="bluecol" style="width: 16%">聯絡人 </td>
                            <td class="whitecol" style="width: 34%">
                                <asp:Label ID="ContactName" runat="server"></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">年度 </td>
                            <td class="whitecol">
                                <asp:Label ID="YearList" runat="server"></asp:Label><asp:Label ID="labAPPSTAGE" runat="server"></asp:Label></td>
                            <td class="bluecol">
                                <asp:Label ID="LabTMID" runat="server">訓練職類</asp:Label></td>
                            <td class="whitecol">
                                <asp:Label ID="TrainText" runat="server"></asp:Label>
                                <asp:Label ID="JobText" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="Labcjob" runat="server">通俗職類</asp:Label></td>
                            <td colspan="3" class="whitecol">
                                <%--<asp:Label ID="CjobNO" runat="server"></asp:Label>--%>
                                <asp:Label ID="CjobName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">變更項目 </td>
                            <td class="whitecol">
                                <asp:Label ID="ChgItem" runat="server"></asp:Label></td>
                            <td class="bluecol">申請變更日 </td>
                            <td class="whitecol">
                                <asp:Label ID="ApplyDate" runat="server"  ></asp:Label></td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級名稱 </td>
                            <td class="whitecol">
                                <asp:Label ID="ClassCName" runat="server"></asp:Label>
                                <asp:Label ID="PointYN" runat="server"></asp:Label>
                            </td>
                             <td class="bluecol">線上送件時間 </td>
                            <td class="whitecol">
                                <asp:Label ID="lbONLINESENDDATE" runat="server" ></asp:Label></td>
                        </tr>
                        <tr id="trlbD20KNAME" runat="server">
                            <td class="bluecol">政府政策性產業 </td>
                            <td colspan="3" class="whitecol">
                                <asp:Label ID="lbD20KNAME" runat="server"></asp:Label>
                                <asp:Label ID="lbD25KNAME" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <div id="divCo128" runat="server">
                                    <table cellpadding="1" cellspacing="1" width="100%" border="0">
                                        <tr>
                                            <td class="bluecol"></td>
                                            <td colspan="3" class="whitecol">
                                                <asp:Label ID="lab_msg_4" runat="server"></asp:Label>
                                                <%--<label id="labmsg_4"></label><asp:Label ID="lab_msg_4" runat="server"></asp:Label>--%>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need" style="width: 16%">申請變更函送日期</td>
                                            <td class="whitecol" style="width: 34%">
                                                <span runat="server">
                                                    <asp:TextBox ID="SENDDATE4" Width="40%" onfocus="this.blur()" runat="server" MaxLength="10"></asp:TextBox>
                                                    <img style="cursor: pointer" onclick="javascript:show_calendar('SENDDATE4','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                                </span>
                                                <asp:Button ID="btnChkSend4" runat="server" Text="檢核" CssClass="asp_Export_M" />
                                            </td>
                                            <td class="bluecol_need" style="width: 16%">申請變更函送狀態(2-1-1)</td>
                                            <td class="whitecol" style="width: 34%">
                                                <asp:RadioButtonList ID="STATUS4" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                    <asp:ListItem Value="1">依規定辦理</asp:ListItem>
                                                    <asp:ListItem Value="2">逾期(扣分)</asp:ListItem>
                                                    <asp:ListItem Value="3">逾期(不扣分)</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol_need">資料內容符合規定(2-1-2)</td>
                                            <td class="whitecol">
                                                <asp:DropDownList ID="ddlISPASS4" runat="server"></asp:DropDownList>
                                                <%--<asp:RadioButtonList ID="ISPASS4" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                    <asp:ListItem Value="R">駁回</asp:ListItem>
                                                    <asp:ListItem Value="Y">符合規定</asp:ListItem>
                                                    <asp:ListItem Value="N">不符合規定</asp:ListItem>
                                                </asp:RadioButtonList>--%>
                                            </td>
                                            <td class="bluecol_need">逾期週數(2-1-1)</td>
                                            <td class="whitecol">
                                                <asp:RadioButtonList ID="OVERWEEK4" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                                    <asp:ListItem Value="1">1週以內</asp:ListItem>
                                                    <asp:ListItem Value="2">1週以上</asp:ListItem>
                                                    <asp:ListItem Value="3">停辦</asp:ListItem>
                                                    <asp:ListItem Value="9">無逾期</asp:ListItem>
                                                </asp:RadioButtonList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td class="bluecol">不納入審查計分變更次數</td>
                                            <td class="whitecol">
                                                <asp:CheckBox ID="chkbox_NOINC4" runat="server" Text="不納入" />
                                            </td>
                                            <td class="bluecol"></td><td class="whitecol"></td>
                                            <%--<td class="bluecol">政策性課程不扣分</td><td class="whitecol"><asp:CheckBox ID="chkbox_NODEDUC4" runat="server" Text="不扣分" /></td>--%>
                                        </tr>
                                        <tr id="tr_SET_EnterDate" runat="server">
                                            <td class="bluecol_need">報名開始日期</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="SEnterDate" Width="40%" MaxLength="10" onfocus="this.blur()" runat="server"></asp:TextBox>
                                                <span id="sp_imgSEnterDate" runat="server">
                                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= SEnterDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                                </span>
                                                <asp:DropDownList ID="SEnterDate_HR" runat="server" AppendDataBoundItems="True"></asp:DropDownList>時：                       
                                                <asp:DropDownList ID="SEnterDate_MI" runat="server" AppendDataBoundItems="True"></asp:DropDownList>分&nbsp;
                                            </td>
                                            <td class="bluecol_need">報名結束日期</td>
                                            <td class="whitecol">
                                                <asp:TextBox ID="FEnterDate" Width="40%" MaxLength="10" onfocus="this.blur()" runat="server"></asp:TextBox>
                                                <span id="sp_imgFEnterDate" runat="server">
                                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= FEnterDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                                </span>
                                                <asp:DropDownList ID="FEnterDate_HR" runat="server" AppendDataBoundItems="True"></asp:DropDownList>時：                       
                                                <asp:DropDownList ID="FEnterDate_MI" runat="server" AppendDataBoundItems="True"></asp:DropDownList>分&nbsp;
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">審核結果 </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="ChkMode" runat="server">
                                    <asp:ListItem Value="">===請選擇===</asp:ListItem>
                                    <asp:ListItem Value="Y">審核通過</asp:ListItem>
                                    <asp:ListItem Value="N">審核不通過</asp:ListItem>
                                </asp:DropDownList>
                                <asp:Button ID="But_Sub" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Button1" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="btnDelete" runat="server" Text="刪除" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="Btn_SAVE2" runat="server" Text="儲存" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                                <asp:RequiredFieldValidator ID="MustChk" runat="server" ErrorMessage="請選擇審核結果" Display="None" ControlToValidate="ChkMode"></asp:RequiredFieldValidator>
                                <%----%>
                            </td>
                        </tr>
                        <tr id="trPACKAGE_DOWNLOAD1" runat="server">
                            <td class="bluecol">檔案下載</td>
                            <td class="whitecol" colspan="3">
                                <%--File package download--%>
                                <asp:Button ID="BTN_PACKAGE_DOWNLOAD1" runat="server" Text="檔案打包下載" CommandName="PACKAGE_DOWNLOAD1" CssClass="asp_Export_M" CausesValidation="False"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <table class="font" id="ReviseTable" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tbody>
                            <tr id="Item1_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <table class="font" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>訓練期間起訖日期： </td>
                                        </tr>
                                        <tr>
                                            <td>自<asp:Label ID="BSDate" runat="server"></asp:Label>
                                                至<asp:Label ID="BEDate" runat="server"></asp:Label>
                                            </td>
                                        </tr>
                                    </table>
                                    <table class="font" id="Table_Sign1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td>報名起訖日期： </td>
                                        </tr>
                                        <tr>
                                            <td>自<asp:Label ID="Old_SEnterDate2" runat="server"></asp:Label>
                                                至<asp:Label ID="Old_FEnterDate2" runat="server"></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>甄試日期： </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:Label ID="Old_Examdate" runat="server"></asp:Label>
                                                <asp:Label ID="Old_ExamPeriod" runat="server"></asp:Label>
                                                <asp:HiddenField ID="HidOld_ExamPeriod" runat="server" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>報到日期： </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:Label ID="Old_CheckInDate" runat="server"></asp:Label></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item1_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <table class="font" id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>訓練期間起訖日期： </td>
                                        </tr>
                                        <tr>
                                            <td>自<asp:Label ID="ASDate" runat="server"></asp:Label>
                                                至<asp:Label ID="AEDate" runat="server"></asp:Label>
                                            </td>
                                        </tr>
                                    </table>
                                    <table class="font" id="Table_Sign2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td>報名起訖日期： </td>
                                        </tr>
                                        <tr>
                                            <td>自<asp:Label ID="New_SEnterDate2" runat="server"></asp:Label>
                                                至<asp:Label ID="New_FEnterDate2" runat="server"></asp:Label>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>甄試日期： </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:Label ID="New_Examdate" runat="server"></asp:Label>
                                                <asp:Label ID="New_ExamPeriod" runat="server"></asp:Label>
                                                <asp:HiddenField ID="HidNew_ExamPeriod" runat="server" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>報到日期： </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:Label ID="New_CheckInDate" runat="server"></asp:Label></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item2_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>日期：<asp:Label ID="TimeSDate" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:CheckBoxList ID="TimeSClass" runat="server" CssClass="font" RepeatColumns="2"></asp:CheckBoxList></td>
                                        </tr>
                                    </table>
                                    <table class="font">
                                        <tr>
                                            <td valign="top">課程： </td>
                                            <td>
                                                <asp:Label ID="EditSClass" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td valign="top">節次： </td>
                                            <td>
                                                <asp:Label ID="EditSClassItem" runat="server"></asp:Label></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item2_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <table class="font" id="Table6" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>日期：<asp:Label ID="TimeEDate" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:CheckBoxList ID="TimeEClass" runat="server" CssClass="font" RepeatColumns="2"></asp:CheckBoxList></td>
                                        </tr>
                                    </table>
                                    <table class="font">
                                        <tr>
                                            <td valign="top">課程： </td>
                                            <td>
                                                <asp:Label ID="EditEClass" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td valign="top">節次： </td>
                                            <td>
                                                <asp:Label ID="EditEClassItem" runat="server"></asp:Label></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item3_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <table class="font" id="Table7" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>日期：<asp:Label ID="PlaceDate" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:CheckBoxList ID="SPlace" runat="server" CssClass="font" RepeatColumns="2"></asp:CheckBoxList></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item3_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>地點：<asp:Label ID="EPlace" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item4_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <table class="font" id="Table8" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>學科：<asp:Label ID="SSumSci" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>一般學科：<asp:Label ID="SGenSci" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>專業學科：<asp:Label ID="SProSci" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>術科：<asp:Label ID="SProTech" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>其他：<asp:Label ID="SOther" runat="server"></asp:Label></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item4_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <table class="font" id="Table9" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>學科：<asp:Label ID="ESumSci" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>一般學科：<asp:Label ID="EGenSci" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>專業學科：<asp:Label ID="EProSci" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>術科：<asp:Label ID="EProTech" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>其他：<asp:Label ID="EOther" runat="server"></asp:Label></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item5_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <table class="font" id="Table10" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>日期：<asp:Label ID="TechDate" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:CheckBoxList ID="STeacher" runat="server" CssClass="font" RepeatColumns="2"></asp:CheckBoxList></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item5_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>師資姓名：&nbsp;<asp:Label ID="OLessonTeah1" runat="server"></asp:Label>
                                    &nbsp;,&nbsp;助教姓名(1)：&nbsp;&nbsp;<asp:Label ID="OLessonTeah2" runat="server"></asp:Label>
                                    &nbsp;,&nbsp;助教姓名(2)：&nbsp;&nbsp;<asp:Label ID="OLessonTeah3" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr id="Item6_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>班別名稱：<asp:Label ID="OClassName" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item6_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>班別名稱：<asp:Label ID="ChangeOClassName" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item7_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <table class="font" id="Table11" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>班別名稱：<asp:Label ID="OClassName2" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>期別：<asp:Label ID="CyclType" runat="server"></asp:Label></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item7_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>期別：<asp:Label ID="ChangeCyclType" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item8_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <asp:Label ID="TAddress1" runat="server"></asp:Label>
                                    <input id="OldData8_1" type="hidden" name="OldData8_1" runat="server">
                                    <input id="OldData8_2" type="hidden" name="OldData8_2" runat="server">
                                    <input id="OldData8_3" type="hidden" name="OldData8_3" runat="server">
                                    <input id="hidOldData8_6W" type="hidden" runat="server" />
                                </td>
                            </tr>
                            <tr id="Item8_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <asp:Label ID="TAddress2" runat="server"></asp:Label>
                                    <input id="NewData8_1" type="hidden" name="NewData8_1" runat="server">
                                    <input id="NewData8_2" type="hidden" name="NewData8_2" runat="server">
                                    <input id="NewData8_3" type="hidden" name="NewData8_3" runat="server">
                                    <input id="hidNewData8_6W" type="hidden" runat="server" />
                                </td>
                            </tr>
                            <tr id="Item9_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <asp:Label ID="NotOpen1" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item9_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <asp:Label ID="NotOpen2" runat="server"></asp:Label><input id="NewData9_1" type="hidden" runat="server" /><input id="OldData9_1" type="hidden" runat="server" /></td>
                            </tr>
                            <tr id="Item10_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <asp:Label ID="TrainTime1" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item10_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <asp:Label ID="TrainTime2" runat="server"></asp:Label><input id="NewData10_1" type="hidden" name="NewData10_1" runat="server"></td>
                            </tr>
                            <tr id="Item11_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <asp:Label ID="TeacherName1" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item11_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <asp:Label ID="TeacherName1_2" runat="server"></asp:Label>
                                    <input id="NewData11_1" type="hidden" name="NewData11_1" runat="server">
                                    <asp:HiddenField ID="Hid_NewData11_3" runat="server" />
                                </td>
                            </tr>
                            <tr id="Item20_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <asp:Label ID="TeacherName2" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item20_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <asp:Label ID="TeacherName2_2" runat="server"></asp:Label>
                                    <input id="NewData20_1" type="hidden" name="NewData20_1" runat="server">
                                    <asp:HiddenField ID="Hid_NewData20_3" runat="server" />
                                </td>
                            </tr>
                            <tr id="Item12_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <asp:Label ID="OldData12_1" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item12_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <asp:Label ID="NewData12_1" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item14_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td class="whitecol">學科場地1／上課地址&nbsp;&nbsp;&nbsp;<%--<asp:Label ID="SciPlaceID" runat="server" Width="380px" BorderStyle="Groove" Visible="False"></asp:Label>--%>
                                    <asp:Label ID="SciPlaceID" runat="server" Width="400px" BorderStyle="Groove"></asp:Label><input id="OldData14_1" type="hidden" name="OldData14_1" runat="server"><br>
                                    學科場地2／上課地址&nbsp;&nbsp;&nbsp;<%--<asp:Label ID="SciPlaceID2" runat="server" Width="380px" BorderStyle="Groove" Visible="False"></asp:Label>--%>
                                    <asp:Label ID="SciPlaceID2" runat="server" Width="400px" BorderStyle="Groove"></asp:Label><input id="OldData14_3" type="hidden" name="OldData14_3" runat="server"><br>
                                    術科場地1／上課地址&nbsp;&nbsp;&nbsp;<%--<asp:Label ID="TechPlaceID" runat="server" Width="380px" BorderStyle="Groove" Visible="False"></asp:Label>--%>
                                    <asp:Label ID="TechPlaceID" runat="server" Width="400px" BorderStyle="Groove"></asp:Label><input id="OldData14_2" type="hidden" name="OldData14_2" runat="server"><br>
                                    術科場地2／上課地址&nbsp;&nbsp;&nbsp;<%--<asp:Label ID="TechPlaceID2" runat="server" Width="380px" BorderStyle="Groove" Visible="False"></asp:Label>--%>
                                    <asp:Label ID="TechPlaceID2" runat="server" Width="400px" BorderStyle="Groove"></asp:Label><input id="OldData14_4" type="hidden" name="OldData14_4" runat="server"><br>
                                    <%-- 學科上課地址&nbsp;<asp:Label ID="AddressSciPTID" runat="server" Width="380px" BorderStyle="Groove"></asp:Label>
                                    <input id="OldData8_4" type="hidden" name="OldData8_4" runat="server" />
                                    <br>術科上課地址&nbsp;<asp:Label ID="AddressTechPTID" runat="server" Width="380px" BorderStyle="Groove"></asp:Label>
                                    <input id="OldData8_5" type="hidden" name="OldData8_5" runat="server" />--%>
                                    <asp:HiddenField ID="Hid_OldData8_4" runat="server" />
                                    <asp:HiddenField ID="Hid_OldData8_5" runat="server" />
                                    <asp:HiddenField ID="Hid_OldData8_6" runat="server" />
                                    <asp:HiddenField ID="Hid_OldData8_7" runat="server" />

                                </td>
                            </tr>
                            <tr id="Item14_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td class="whitecol">學科場地1／上課地址&nbsp;&nbsp;&nbsp;&nbsp;
                                    <asp:DropDownList ID="NewData14_1" runat="server" Enabled="False"></asp:DropDownList>
                                    <br>
                                    學科場地2／上課地址&nbsp;&nbsp;&nbsp;&nbsp;
                                    <asp:DropDownList ID="NewData14_3" runat="server" Enabled="False"></asp:DropDownList>
                                    <br>
                                    術科場地1／上課地址&nbsp;&nbsp;&nbsp;&nbsp;
                                    <asp:DropDownList ID="NewData14_2" runat="server" Enabled="False"></asp:DropDownList>
                                    <br>
                                    術科場地2／上課地址&nbsp;&nbsp;&nbsp;&nbsp;
                                    <asp:DropDownList ID="NewData14_4" runat="server" Enabled="False"></asp:DropDownList>
                                    <asp:HiddenField ID="Hid_NewData8_4" runat="server" />
                                    <asp:HiddenField ID="Hid_NewData8_5" runat="server" />
                                    <asp:HiddenField ID="Hid_NewData8_6" runat="server" />
                                    <asp:HiddenField ID="Hid_NewData8_7" runat="server" />
                                    <%--<br>
                                    學科上課地址&nbsp;
                                    <asp:DropDownList ID="NewData8_4" runat="server" Enabled="False"></asp:DropDownList>
                                    <asp:HiddenField ID="Hid_NewData8_4" runat="server" />
                                    <br>
                                    術科上課地址&nbsp;
                                    <asp:DropDownList ID="NewData8_5" runat="server" Enabled="False"></asp:DropDownList><br>
                                    <asp:HiddenField ID="Hid_NewData8_5" runat="server" />--%>
                                </td>
                            </tr>
                            <tr id="Item15_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="星期">
                                                <HeaderStyle Width="20%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="OldWeeks1" runat="server"></asp:Label>
                                                </ItemTemplate>
                                                <EditItemTemplate>
                                                    <asp:DropDownList ID="OldWeeks2" runat="server">
                                                    </asp:DropDownList>
                                                </EditItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="上課時段">
                                                <HeaderStyle Width="80%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="OldTimes1" runat="server"></asp:Label>
                                                </ItemTemplate>
                                                <EditItemTemplate>
                                                    <asp:TextBox ID="OldTimes2" runat="server"></asp:TextBox>
                                                </EditItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr id="Item15_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="星期">
                                                <HeaderStyle Width="20%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="NewWeeks1" runat="server"></asp:Label>
                                                </ItemTemplate>
                                                <EditItemTemplate>
                                                    <asp:DropDownList ID="NewWeeks2" runat="server">
                                                    </asp:DropDownList>
                                                </EditItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="上課時段">
                                                <HeaderStyle Width="80%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="NewTimes1" runat="server"></asp:Label>
                                                </ItemTemplate>
                                                <EditItemTemplate>
                                                    <asp:TextBox ID="NewTimes2" runat="server"></asp:TextBox>
                                                </EditItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr id="Item16_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容 </td>
                                <td>
                                    <asp:Label ID="OldData16_1" runat="server"></asp:Label></td>
                            </tr>
                            <tr id="Item16_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td>
                                    <asp:Label ID="NewData16_1" runat="server"></asp:Label></td>
                            </tr>
                            <!--20080825 andy add 報名起迄日 Start-->
                            <tr id="Item17_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容</td>
                                <td>
                                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>報名起訖日期： </td>
                                        </tr>
                                        <tr>
                                            <td>自<asp:Label ID="Old_SEnterDate" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>至<asp:Label ID="Old_FEnterDate" runat="server"></asp:Label></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item17_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td class="whitecol">
                                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>報名起訖日期： </td>
                                        </tr>
                                        <tr>
                                            <td>自<asp:TextBox ID="New_SEnterDate" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox></td>
                                        </tr>
                                        <tr>
                                            <td>至<asp:TextBox ID="New_FEnterDate" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item19_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容</td>
                                <td>
                                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>包班種類：
											    <asp:Label ID="PackageTypeOld" runat="server"></asp:Label>
                                                <input id="hidPackageTypeOld" type="hidden" name="hidPackageTypeOld" runat="server">
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>包班事業單位：
											<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td>
                                                        <asp:DataGrid ID="DG_BusPackageOld" runat="server" Width="100%" CssClass="font" CellPadding="8">
                                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" CssClass="head_navy"></HeaderStyle>
                                                            <PagerStyle Visible="False"></PagerStyle>
                                                        </asp:DataGrid>
                                                    </td>
                                                </tr>
                                            </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item19_2" runat="server">
                                <td class="bluecol">變更內容</td>
                                <td>
                                    <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>包班種類：<asp:Label ID="PackageTypeNew" runat="server"></asp:Label>
                                                <input id="hidPackageTypeNew" type="hidden" name="hidPackageTypeNew" runat="server"></td>
                                        </tr>
                                        <tr>
                                            <td>包班事業單位：
											<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td>
                                                        <asp:DataGrid ID="DG_BusPackageNew" runat="server" Width="100%" CssClass="font" CellPadding="8">
                                                            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" CssClass="head_navy"></HeaderStyle>
                                                            <PagerStyle Visible="False"></PagerStyle>
                                                        </asp:DataGrid>
                                                    </td>
                                                </tr>
                                            </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item21_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容</td>
                                <td>
                                    <table id="tbDataGrid21Old" runat="server" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td class="table_title">
                                                <asp:Label ID="labcost21txt1Old" runat="server" Text="費用列表"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid21Old_1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="項目">
                                                            <HeaderStyle Width="60%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="CostName" runat="server">CostName</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="單價">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:HiddenField ID="HidCostID" runat="server" />
                                                                <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="數量">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="計價單位">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="ItemCost" runat="server">ItemCost</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="小計">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                                <asp:DataGrid ID="DataGrid21Old_2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="單價">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="人數">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="時數">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="ItemCost" runat="server">ItemCost</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="小計">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                                <asp:DataGrid ID="DataGrid21Old_3" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="單價">
                                                            <HeaderStyle Width="33%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="人數">
                                                            <HeaderStyle Width="33%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="小計">
                                                            <HeaderStyle Width="34%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                                <asp:DataGrid ID="DataGrid21Old_4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="項目">
                                                            <HeaderStyle Width="70%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="CostName" runat="server">CostName</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="單價">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="數量">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="小計">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                                    <tr id="AdmGrantTROld" runat="server">
                                                        <td class="bluecol" width="20%">行政管理費 </td>
                                                        <td class="whitecol">
                                                            <asp:Label ID="AdmCostOld" runat="server"></asp:Label></td>
                                                    </tr>
                                                    <tr id="TaxGrantTROld" runat="server">
                                                        <td class="bluecol">營業稅 </td>
                                                        <td class="whitecol">
                                                            <asp:Label ID="TaxCostOld" runat="server"></asp:Label></td>
                                                    </tr>
                                                    <tr id="trTotalCost1Old" runat="server">
                                                        <td class="bluecol">總計 </td>
                                                        <td class="whitecol">
                                                            <asp:Label ID="TotalCost1Old" runat="server"></asp:Label>
                                                            <input id="HidAdmGrantOld" type="hidden" value="0" runat="server">
                                                            <input id="HidTaxGrantOld" type="hidden" value="0" runat="server">
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item21_2" runat="server">
                                <td class="bluecol">變更內容</td>
                                <td>
                                    <table id="tbDataGrid21New" runat="server" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td class="table_title">
                                                <asp:Label ID="labcost21txt1New" runat="server" Text="費用列表"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="DataGrid21New_1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="項目">
                                                            <HeaderStyle Width="50%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="CostName" runat="server">CostName</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="單價">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:HiddenField ID="HidCostID" runat="server" />
                                                                <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="數量">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="計價單位">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="ItemCost" runat="server">ItemCost</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="小計">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Button ID="btnDel1" runat="server" CausesValidation="False" Text="刪除" CommandName="btnDel1" CssClass="asp_button_M"></asp:Button>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                                <asp:DataGrid ID="DataGrid21New_2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="單價">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="人數">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="時數">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="ItemCost" runat="server">ItemCost</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="小計">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Button ID="btnDel1" runat="server" CausesValidation="False" Text="刪除" CommandName="btnDel1" CssClass="asp_button_M"></asp:Button>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                                <asp:DataGrid ID="DataGrid21New_3" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="單價">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="人數">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="小計">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <HeaderStyle Width="25%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Button ID="btnDel1" runat="server" CausesValidation="False" Text="刪除" CommandName="btnDel1" CssClass="asp_button_M"></asp:Button>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                                <asp:DataGrid ID="DataGrid21New_4" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <Columns>
                                                        <asp:TemplateColumn HeaderText="項目">
                                                            <HeaderStyle Width="60%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="CostName" runat="server">CostName</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="單價">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="OPrice" runat="server">OPrice</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="數量">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="Itemage" runat="server">Itemage</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="小計">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="subtotal" runat="server">subtotal</asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <HeaderStyle Width="10%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Button ID="btnDel1" runat="server" CausesValidation="False" Text="刪除" CommandName="btnDel1" CssClass="asp_button_M"></asp:Button>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <table cellspacing="1" cellpadding="1" width="100%" border="0">
                                                    <tr id="AdmGrantTRNew" runat="server">
                                                        <td class="bluecol">行政管理費 </td>
                                                        <td class="whitecol">
                                                            <asp:Label ID="AdmCostNew" runat="server"></asp:Label></td>
                                                    </tr>
                                                    <tr id="TaxGrantTRNew" runat="server">
                                                        <td class="bluecol">營業稅 </td>
                                                        <td class="whitecol">
                                                            <asp:Label ID="TaxCostNew" runat="server"></asp:Label></td>
                                                    </tr>
                                                    <tr id="trTotalCost1New" runat="server">
                                                        <td class="bluecol">總計 </td>
                                                        <td class="whitecol">
                                                            <asp:Label ID="TotalCost1New" runat="server"></asp:Label>
                                                            <input id="HidAdmGrantNew" type="hidden" value="0" runat="server">
                                                            <input id="HidTaxGrantNew" type="hidden" value="0" runat="server">
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr id="Item22_1" runat="server">
                                <td class="bluecol" style="width: 16%">原計畫內容</td>
                                <td class="whitecol">
                                    <asp:HiddenField ID="Hid_DISTANCE" runat="server" />
                                    <asp:Label ID="lab_DISTANCE" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr id="Item22_2" runat="server">
                                <td class="bluecol">變更內容 </td>
                                <td class="whitecol">
                                    <asp:HiddenField ID="Hid_DISTANCE_new" runat="server" />
                                    <asp:Label ID="lab_DISTANCE_new" runat="server" Text=""></asp:Label>
                                </td>
                            </tr>
                            <!-- end -->
                            <tr>
                                <td class="bluecol" style="width: 16%">變更原因 </td>
                                <td style="width: 84%">
                                    <asp:Label ID="changeReason" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">變更說明 </td>
                                <td>
                                    <asp:Label ID="ReviseReason" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">審核說明 </td>
                                <td>
                                    <asp:TextBox ID="ReviseCont" runat="server" Width="77%" TextMode="MultiLine" Rows="5"></asp:TextBox></td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
            <tr id="Item11_3" runat="server">
                <td style="width: 100%">
                    <div>
                        <table class="font" id="tbDataGrid21" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td class="table_title" align="center">授課教師 </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid21" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="8%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="seqno" runat="server"></asp:Label>
                                                    <input id="HidTechID" runat="server" type="hidden" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="教師姓名">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="14%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="TeachCName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="學歷">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="18%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="DegreeName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="專業領域">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="Specialty1" runat="server"></asp:Label>
                                                    <%--<asp:Label ID="ProLicense" runat="server"></asp:Label>--%>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="遴選辦法說明">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:TextBox ID="TeacherDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                                                    <input id="btn_TCTYPEA" type="button" value="..." runat="server" class="button_b_Mini" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
            <tr id="Item20_3" runat="server">
                <td style="width: 100%">
                    <div>
                        <table class="font" id="tbDataGrid22" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td align="center" class="table_title">授課助教 </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid22" runat="server" Width="100%" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="序號">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="8%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="seqno" runat="server"></asp:Label>
                                                    <input id="HidTechID" runat="server" type="hidden" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="助教姓名">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="14%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="TeachCName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="學歷">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="18%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="DegreeName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="專業領域">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" Width="20%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="Specialty1" runat="server"></asp:Label>
                                                    <%--<asp:Label ID="ProLicense" runat="server"></asp:Label>--%>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="遴選辦法說明">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:TextBox ID="TeacherDesc" runat="server" Width="80%" Height="80px" TextMode="MultiLine"></asp:TextBox>
                                                    <input id="btn_TCTYPEB" type="button" value="..." runat="server" class="button_b_Mini" />
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>

        </table>
        <!--- DG4  Start--->
        <table id="dg4_table" width="100%">
            <tr>
                <td id="dg4_dt" runat="server" class="table_title">課程表申請變更後</td>
                <td id="dg3_dt" runat="server" class="table_title">課程表申請變更前</td>
            </tr>
            <tr>
                <td>
                    <asp:DataGrid ID="Datagrid4" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                        <ItemStyle></ItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="日期">
                                <ItemTemplate>
                                    <asp:Label ID="lab_STrainDate" runat="server"></asp:Label>
                                    <input id="hide_ID1" type="hidden" runat="server" />
                                    <input id="hide_PTDRID" type="hidden" runat="server" />
                                    <input id="hide_PTDID" type="hidden" runat="server" />
                                </ItemTemplate>
                                <%--<EditItemTemplate>
                                    &nbsp;<asp:Button ID="Button12" Style="display: none" runat="server"></asp:Button>
                                </EditItemTemplate>--%>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="授課時段">
                                <ItemStyle Wrap="false"></ItemStyle>
                                <ItemTemplate>
                                    <asp:CheckBox ID="OldTPERIOD28_1t" runat="server" Text="早上" ToolTip="7:00-13:00" Enabled="false" /><br />
                                    <asp:CheckBox ID="OldTPERIOD28_2t" runat="server" Text="下午" ToolTip="13:00-18:00" Enabled="false" /><br />
                                    <asp:CheckBox ID="OldTPERIOD28_3t" runat="server" Text="晚上" ToolTip="18:00-22:00" Enabled="false" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="授課時間">
                                <ItemTemplate>
                                    <asp:Label ID="lab_PName" Width="80%" runat="server"></asp:Label>
                                    <input id="hide_ID2" type="hidden" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="時數">
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:Label ID="lab_PHour" runat="server"></asp:Label><input id="hide_ID3" type="hidden" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="技檢訓練時數">
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:Label ID="lab_EHour" runat="server"></asp:Label><input id="hide_ID9" type="hidden" runat="server" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="課程進度／內容">
                                <ItemTemplate>
                                    <asp:TextBox ID="txt_PCont" runat="server" Width="150px" Enabled="False" Height="70px" TextMode="MultiLine" Rows="5"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="學／術科">
                                <ItemStyle CssClass="whitecol" />
                                <ItemTemplate>
                                    <asp:DropDownList ID="list_Classification" runat="server" Enabled="False">
                                        <asp:ListItem Value="1">學科</asp:ListItem>
                                        <asp:ListItem Value="2">術科</asp:ListItem>
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="上課地點">
                                <HeaderStyle></HeaderStyle>
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" Wrap="false" />
                                <ItemTemplate>
                                    <asp:Label ID="lab_PTID" runat="server"></asp:Label>
                                    <input id="hide_PTID" type="hidden" runat="server" />
                                    <%--<input id="hide_ID4" type="hidden" runat="server" />--%>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="遠距教學">
                                <ItemTemplate>
                                    <asp:CheckBox ID="bx_FARLEARN" runat="server" Enabled="false" />
                                    <input id="hide_FARLEARN" type="hidden" runat="server" />
                                    <%--<input id="hide_ID8" type="hidden" runat="server" />--%>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="任課教師">
                                <HeaderStyle></HeaderStyle>
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <%--<asp:Label ID="lab_TechID" runat="server"></asp:Label>--%>
                                    <asp:TextBox ID="newTechText" runat="server" Width="70px" Enabled="False"></asp:TextBox>
                                    <input id="hide_TechID" type="hidden" runat="server" />
                                    <%--<input id="hide_ID5" type="hidden" runat="server" />--%>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="助教">
                                <HeaderStyle></HeaderStyle>
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <%--<asp:Label ID="lab_TechID2" runat="server"></asp:Label>--%>
                                    <asp:TextBox ID="newTech2Text" runat="server" Width="70px" Enabled="False"></asp:TextBox>
                                    <input id="hide_TechID2" type="hidden" runat="server" />
                                    <%--<input id="hide_ID6" type="hidden" runat="server" />--%>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
                <td>
                    <asp:DataGrid ID="Datagrid3" runat="server" Width="100%" AutoGenerateColumns="False" CellPadding="8">
                        <ItemStyle></ItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <Columns>
                            <asp:TemplateColumn HeaderText="日期">
                                <ItemTemplate>
                                    <asp:Label ID="OldSTrainDateLabel" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="授課時段">
                                <ItemStyle Wrap="false"></ItemStyle>
                                <ItemTemplate>
                                    <asp:CheckBox ID="OldTPERIOD28_1t" runat="server" Text="早上" ToolTip="7:00-13:00" Enabled="false" /><br />
                                    <asp:CheckBox ID="OldTPERIOD28_2t" runat="server" Text="下午" ToolTip="13:00-18:00" Enabled="false" /><br />
                                    <asp:CheckBox ID="OldTPERIOD28_3t" runat="server" Text="晚上" ToolTip="18:00-22:00" Enabled="false" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="授課時間">
                                <ItemTemplate>
                                    <asp:Label ID="OldPNameLabel" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="時數">
                                <%--<HeaderStyle Wrap="False" Width="6%"></HeaderStyle>--%>
                                <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:Label ID="OldPHourLabel" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="技檢訓練時數">
                                <%--<HeaderStyle Wrap="False" Width="6%"></HeaderStyle>--%>
                                <ItemStyle Width="6%" Wrap="false" HorizontalAlign="Center" />
                                <ItemStyle HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <asp:Label ID="OldEHourLabel" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="課程進度／內容">
                                <ItemTemplate>
                                    <asp:TextBox ID="OldPContText" runat="server" Width="150px" Rows="5" TextMode="MultiLine" Height="70px" Enabled="False"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="學／術科">
                                <ItemTemplate>
                                    <asp:DropDownList ID="OlddrpClassification1" runat="server" Enabled="False">
                                        <asp:ListItem Value="1">學科</asp:ListItem>
                                        <asp:ListItem Value="2">術科</asp:ListItem>
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="上課地點">
                                <ItemTemplate>
                                    <asp:DropDownList ID="OlddrpPTID" runat="server" Enabled="False">
                                    </asp:DropDownList>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="遠距教學">
                                <ItemTemplate>
                                    <asp:CheckBox ID="OldFARLEARN" runat="server" Enabled="false" />
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="任課教師">
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <input id="OldTech1Value" type="hidden" size="3" name="OldTech1Value" runat="server">
                                    <asp:TextBox ID="OldTech1Text" runat="server" Width="70px" Enabled="False"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:TemplateColumn HeaderText="助教">
                                <ItemStyle CssClass="whitecol" HorizontalAlign="Center" />
                                <ItemTemplate>
                                    <input id="OldTech2Value" type="hidden" size="3" name="OldTech2Value" runat="server">
                                    <asp:TextBox ID="OldTech2Text" runat="server" Width="70px" Enabled="False"></asp:TextBox>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                    </asp:DataGrid>
                </td>
            </tr>
        </table>
        <!--- DG4  E-->
        <%--<div>
        </div>--%>
        <input id="hid_NewData12" type="hidden" runat="server" />
        <asp:ValidationSummary ID="TotalMsg" runat="server" DisplayMode="List" ShowMessageBox="True" ShowSummary="False"></asp:ValidationSummary>
        <input id="hid_AltDataID" type="hidden" name="hid_AltDataID" runat="server" />
        <input id="hid_TMID" type="hidden" name="hid_TMID" runat="server" />
        <input id="hid_USE_PLAN_REVISESUB" type="hidden" runat="server" />

        <asp:HiddenField ID="Hid_ComIDNO" runat="server" />
        <asp:HiddenField ID="Hid_RID1" runat="server" />
        <asp:HiddenField ID="Hid_sFENTERDATE2" runat="server" />
        <asp:HiddenField ID="Hid_PlanKind" runat="server" />
        <asp:HiddenField ID="Hid_CostMode" runat="server" />
        <asp:HiddenField ID="Hid_AdmPercent" runat="server" />
        <asp:HiddenField ID="Hid_TaxPercent" runat="server" />
        <asp:HiddenField ID="Hid_OnShellDate" runat="server" />
        <asp:HiddenField ID="Hid_stdate" runat="server" />
        <asp:HiddenField ID="Hid_stdate_7" runat="server" />
        <asp:HiddenField ID="Hid_ApplyDate" runat="server" />
        <asp:HiddenField ID="hid_SP_ZIPCODE" runat="server" />
        <asp:HiddenField ID="hid_SP_ZIP6W" runat="server" />
        <asp:HiddenField ID="hid_SP_ADDRESS" runat="server" />
        <asp:HiddenField ID="hid_TP_ZIPCODE" runat="server" />
        <asp:HiddenField ID="hid_TP_ZIP6W" runat="server" />
        <asp:HiddenField ID="hid_TP_ADDRESS" runat="server" />

        <asp:HiddenField ID="hid_OldData2_2" runat="server" />
        <asp:HiddenField ID="hid_OldData2_3" runat="server" />
        <asp:HiddenField ID="hid_NewData2_2" runat="server" />
        <asp:HiddenField ID="hid_NewData2_3" runat="server" />
        <asp:HiddenField ID="hid_OldData3_3" runat="server" />
        <asp:HiddenField ID="hid_NewData3_1" runat="server" />
        <asp:HiddenField ID="hid_OldData5_3" runat="server" />
        <asp:HiddenField ID="hid_NewData5_1" runat="server" />
    </form>
</body>
</html>
