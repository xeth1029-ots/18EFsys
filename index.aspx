<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="index.aspx.vb" Inherits="WDAIIP.index" MasterPageFile="~/LayoutNoHeader.Master" %>

<asp:Content ContentPlaceHolderID="ContentHeader" runat="server">
    <div class="header clearfix">
        <div class="mask-on-jobs"></div>
        <div class="mask-on-jobs-bg"></div>
        <%--<div class="mask-on-jobs-bg"></div>--%>
        <%--style="height: 80px; margin-left: -6px; background-image: url('/images/logo-index-bg.png'); background-repeat: repeat-x;"--%>
        <div class="logo-index col-md-6">
            <a href="<%=Request.ApplicationPath %>main2" target="mainFrame" style="margin: 0px; padding: 0px;">
                <img style="margin: 0px; padding: 0px;" src="/images/wda_logo_on_job.svg" alt="勞動部勞動力發展署-在職訓練資訊管理系統" />
            </a>
        </div>
        <%--style="background-image: url('/images/logo-index-bg.png'); background-repeat: repeat-x;"--%>
        <div class="col-md-6">
            <div class="user-set">
                <div class="user-name">
                    <asp:Label ID="labID" runat="server"></asp:Label>
                </div>
                <div class="user-type">
                    <asp:Label ID="labPlan" runat="server"></asp:Label>
                </div>
            </div>
            <div class="data-set">
                <div class="data-timer">
                    <img src="/images/icon-timer.svg" alt="時鐘圖示" />
                    倒數 <span class="data-timer-number">00:00</span>
                    <button type="submit" class="btn" style="margin-left: 10px; margin-top: -22px; background: #ffffff;" onclick="resetTimer();" title="重設倒數計時"><i class="fa fa-refresh" aria-hidden="true"></i></button>
                </div>
                <div class="data-cal">
                    <span class="data-cal-number-small"><%= Right("0" + (DateTime.Now.Year - 1911).ToString(), 3) %></span><span style="font-size: 9pt;">年</span>
                    <span class="data-cal-number-small"><%= Right("0" + (DateTime.Now.Month).ToString(), 2) %></span><span style="font-size: 9pt;">月</span>
                    <span class="data-cal-number-small"><%= Right("0" + (DateTime.Now.Day).ToString(), 2) %></span><span style="font-size: 9pt;">日</span>
                </div>
            </div>
        </div>
        <%--<table class="fontmsn" id="eMeng" style="visibility: hidden; border-right: #455690 1px solid; border-top: #a6b4cf 1px solid; z-index: 99999; left: 0px; border-left: #a6b4cf 1px solid; border-bottom: #455690 1px solid; position: absolute; top: 0px; height: 100px; background-color: #c9d3f3" cellspacing="1" cellpadding="1" width="180px" border="0" runat="server">--%>
        <%--<div style="float: right;">
            <table class="fontmsn" id="eMeng" style="margin-left: -200px; margin-top: 10px; border-right: #455690 1px solid; border-top: #a6b4cf 1px solid; z-index: 99999; border-left: #a6b4cf 1px solid; border-bottom: #455690 1px solid; position: absolute; top: 0px; height: 100px; background-color: #c9d3f3" cellspacing="1" cellpadding="1" width="180px" border="0" runat="server">
                <tr>
                    <td style="border-right: #b9c9ef 1px solid; padding-right: 10px; border-top: #728eb8 1px solid; padding-left: 10px; font-size: 12px; padding-bottom: 0px; border-left: #728eb8 1px solid; width: 100%; color: #1f336b; padding-top: 5px; border-bottom: #b9c9ef 1px solid;" align="left" background="./images/MsnBack.gif" colspan="1" height="80"><font style="color: red">您閒置系統已超過１５分鐘<br />
                        若未進行資料儲存或點選其他功能，再過５分鐘系統將會自動中斷連線。 </font></td>
                </tr>
            </table>
        </div>--%>
    </div>
</asp:Content>

<asp:Content ContentPlaceHolderID="MainCPH" runat="server">
    <style type="text/css">
        #mainFrame { width: 100%; height: 1024px; border: 0; overflow-x: auto; overflow-y: scroll; }
    </style>
    <div class="col-xs-2">
        <!-- Sidebar Left Menu -->
        <div class="navbg">
            <div class="navigation">
                <h4 class="menu-header">
                    <img src="/images/icon-arrow.svg" alt="功能選單" />功能選單</h4>

                <%-- 動態選單內容 --%>
                <ul id="ulMenu" runat="server"></ul>
                <%-- 靜態選單內容(沒有用) --%>
            </div>
        </div>
        <!-- /#sidebar-wrapper -->
    </div>
    <div class="col-xs-10">
        <div class="type-bar clearfix">
            <h3 class="type-title">
                <img src="/images/icon-document.svg" alt="" />
                <span id="titlePath1"></span>
                <span id="titlePath2"></span>
                <button type="button" class="btn btn-default btn-group-right" onclick="location.href='<%=Request.ApplicationPath %>MOICA_login';" style="margin-left: 5px;"><i class="fa fa-sign-out" aria-hidden="true"></i>登出</button>
                <button type="button" class="btn btn-warning btn-group-right" onclick="return open_help_1();" style="margin-left: 5px;"><i class="fa fa-question-circle"></i>線上說明</button>
                <button type="button" class="btn btn-warning btn-group-right" onclick="changePlan()" style="margin-left: 5px;"><i class="fa fa-window-restore" aria-hidden="true"></i>切換計畫</button>
                <button type="button" class="btn btn-info btn-group-right" onclick="menuClick('sch', '0')" style="margin-left: 5px;"><i class="fa fa-search" aria-hidden="true"></i>功能搜尋</button>
                <button type="button" class="btn btn-info btn-group-right" onclick="show_userInfo()" style="margin-left: 5px;"><i class="fa fa-address-card" aria-hidden="true"></i>帳號資訊</button>
            </h3>
        </div>
        <div class="wrap">
            <!-- page body start -->
            <iframe name="mainFrame" id="mainFrame" src="main2"></iframe>
            <!-- page body end -->
        </div>
    </div>
    <!-- footer start -->
    <div class="footer text-center">
        <div>&nbsp;</div>
        <ul class="list-unstyled">
            <li><b>勞動部勞動力發展署 在職訓練資訊管理系統</b></li>
            <li><b>網頁瀏覽器建議使用&nbsp;Edge、Chrome、Firefox&nbsp;，最佳瀏覽解析度為&nbsp;1024x768&nbsp;以上</b></li>
            <li><b>勞動部勞動力發展署&nbsp;&nbsp;&nbsp;版權所有&nbsp;&nbsp;翻印必究</b></li>
        </ul>
        <%--<div>&nbsp;</div><asp:HiddenField ID="Hid_idx99_tplanid" runat="server" />--%>
        <div>&nbsp;<asp:Label ID="Labmsg1" runat="server" Text="Labx" ForeColor="#D5EEFF"></asp:Label></div>
    </div>
    <!-- footer end -->
</asp:Content>

<asp:Content ContentPlaceHolderID="ContentAUX" runat="server">
    <!-- popDialog() 所需的 container -->
    <div class="modal fade common-dialog" id="commonDialog" role="dialog" aria-labelledby="commonDialogTitle" aria-hidden="true">
        <div class="modal-dialog">
            <div class="modal-content">
                <div class="modal-header bg-primary">
                    <button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">&times;</span><span class="sr-only">Close</span></button>
                    <h3 class="modal-title" id="commonDialogTitle">系統對話框</h3>
                </div>
                <div class="modal-body">
                </div>
                <div class="modal-footer center" style="cursor: pointer">
                    <a class="tablebtn link-pointer btn-close">關閉取消</a>
                </div>
            </div>
        </div>
    </div>
    <script type="text/javascript">
        var tExpire = 1200; //900=15*60 //1200=20*60 // 30(sec);//330(sec)test
        var tCounter = tExpire;
        var tM = '--';
        var tS = '--';
        var tMS = '--';
        var timeoutMsg = "您的登入 Session 已過期, 請重新登入!";
        //$t_eMeng = $('#eMeng');
        $s_lastM = '05:00'; //'19:40';'00:10';'05:00';//最後倒數時間(分鐘)
        //$s_lastM = '00:20'; //'19:40';'00:10';'05:00';//最後倒數時間(分鐘)

        function showMsg2() {
            //debugTrace("Start showMsg2"); //alert(msg2);
            var msg2 = '您閒置系統已超過１５分鐘\n若未進行資料儲存或點選其他功能，再過５分鐘系統將會自動中斷連線。';
            blockAlert(msg2);
        }
        function refreshTimer() {
            $("span.data-timer-number").text(tM + ":" + tS);
        }
        function timerTick() {
            if (tCounter <= 0) {
                blockAlert(timeoutMsg, "逾時", function () {
                    window.location.href = "<%=Request.ApplicationPath %>MOICA_login";
                });
                return;
            }

            var timer = setTimeout(function () { timerTick(); }, 1000);

            tCounter--;
            var m = Math.trunc(tCounter / 60);
            var s = tCounter - (m * 60);
            tM = ((m < 10) ? '0' : '') + m;
            tS = ((s < 10) ? '0' : '') + s;
            tMS = tM + ':' + tS;
            //show_secs,$s_dgmsg1 = "tCounter: " + tCounter + ",tMS : " + tMS; debugTrace($s_dgmsg1);
            if (tMS == $s_lastM) { showMsg2(); }

            refreshTimer();
        }
        function resetTimer() {
            tCounter = tExpire;

            <%-- 
            對主機端發送 Request 以更新 Session
            Request 功能不重要, 所以借用 SelectPlan.aspx 的功能 
            --%>
            var parms = {
                "OP": "Ajax",
                "YR": "2018"
            };
            var url = "<%=Request.ApplicationPath%>SelectPlan.aspx";
            ajaxLoadMore(url, parms, function (resp) {
                if (resp != undefined) {
                    debugTrace("resetTimer success");
                }
            });
        }
        function changePlan() {
            var url = "<%=Request.ApplicationPath %>SelectPlan";
            window.location.href = url;
        }
        function show_userInfo() {
            var url = "<%=Request.ApplicationPath %>eShwUsrInfo";
            window.location.href = url;
        }
        function menuClick(funUrl, funId) {
            hideAllMenu();

            var url = funUrl;

            if (url.indexOf("?") > -1) {
                url += "&";
            } else {
                url += "?";
            }
            url += "ID=" + funId;

            blockUI();
            var mask = $('#loadingMask');  /* === by:20180827 === */
            mask.find("img").show();  /* === by:20180827 === */

            var contentWin = document.getElementById('mainFrame').contentWindow;
            contentWin.location = url;
        }

        var contentDoc;  <%-- iframe content document --%>
        var curSubmitBtnId;

        $(document).ready(function () {

            timerTick();

            $(".navigation ul li a").not(".func").on("click", function () {
                subMenuShow(this);
            });

            $(".menu-header").on('click', function () {
                hideAllMenu();
            });

            // Binding mainFrame OnLoad event
            // 動態 注入(增加) javascript 的引用
            $('#mainFrame').on("load", function () {
                var appPath = "<%=Request.ApplicationPath%>";
                if (appPath != "/") {
                    appPath += "/";
                }
                var contentWin = document.getElementById('mainFrame').contentWindow;
                var contentUrl = contentWin.location.href;
                contentUrl = contentUrl.substring(contentUrl.indexOf("/", 10));
                if (contentUrl.indexOf(appPath) == 0) {
                    contentUrl = contentUrl.substring(appPath.length - 1);
                }
                debugTrace("mainFrame load: " + contentUrl)

                // 將 /Scripts/jquery-3.7.1.min.js 動態加到 頁面中
                var scriptJQuery = contentWin.document.createElement("script");
                scriptJQuery.type = "text/javascript";
                scriptJQuery.src = appPath + "Scripts/jquery-3.7.1.min.js";
                contentWin.document.body.appendChild(scriptJQuery);

                // contentWin 中的 document 物件
                contentDoc = $(contentWin.document);

                // 將 contentWin 中的程式 路徑標題(#FrameTable #TitleLab2)
                // 移至新版面上方功能路徑 bar
                var title1 = "", title2 = "";
                var frameTable = contentDoc.find("#FrameTable");
                if (frameTable != undefined) {
                    title1 = frameTable.find("#TitleLab1");
                    title2 = frameTable.find("#TitleLab2");
                    debugTrace(frameTable);
                    debugTrace(title1);
                    debugTrace(title2);

                    if (title1 != undefined || title2 != undefined) {
                        var typeTitle = $("h3.type-title");
                        typeTitle.find("#titlePath1").html(title1 != undefined ? title1.html() : "");
                        typeTitle.find("#titlePath2").html(title2 != undefined ? title2.html() : "");

                        ((title1 != undefined) ? title1 : title2)
                            .closest("table")
                            .remove(); // 移除 ContentWin 中的功能路徑 table
                    }

                    //處理「首頁-功能路徑」會殘留先前路徑問題，by:20180829
                    var myIndex1 = contentUrl.indexOf("main2");
                    var myIndex2 = contentUrl.indexOf("main2_detail");
                    if (myIndex1 > -1 || myIndex2 > -1) {
                        var typeTitle = $("h3.type-title");
                        typeTitle.find("#titlePath1").html("");
                        typeTitle.find("#titlePath2").html("首頁");
                    }
                }

                // 對所有 contentWin 中的 submit (class 包含 asp_Export_M) button,植入 onclick() 以觸發 unblockUI timer
                contentDoc.find("input[type=submit].asp_Export_M")
                    .on("click", function () {
                        debugTrace(this);
                        curSubmitBtnId = $(this).attr("id");
                        // 匯出 excel 之類的動作, 在contentDoc不會有 onload 事件, 啟動 timer 去 unblockUI()
                        chkCount = 0;
                        chkReadyState();
                    });

                // 對所有 contentWin 中的 form, 植入 onSubmit() 
                // 以加上 submit 時的 blockUI() 效果
                contentDoc.find("form").submit(function (e) {
                    debugTrace(e);

                    blockUI();
                    var mask = $('#loadingMask');  /* === by:20180827 === */
                    mask.find("img").show();  /* === by:20180827 === */

                    <%-- 2018.09.10, 
                     * 因為 originalEvent.explicitOriginalTarget 只有 FireFox 有定義
                     * 不再用下面這一段  

                    // 依觸發按鈕名稱(value值), 判斷是否為匯出類功能
                    // 若存在其他可能名稱要加在這裡
                    //debugTrace("submit btn: " + e.originalEvent.explicitOriginalTarget.value)
                    var submitBtn = e.originalEvent.explicitOriginalTarget;

                    if (submitBtn && submitBtn.value && submitBtn.value.indexOf("匯出") > -1) {
                        // 匯出 excel 之類的動作, 不會有 load 事件, 
                        // 設定 timer 去 unblockUI()
                        chkCount = 0;
                        chkReadyState();
                    }
                    if (submitBtn && $(submitBtn).hasClass("asp_Export_M")) {
                        // 匯出 excel 之類的動作, 不會有 load 事件, 
                        // 設定 timer 去 unblockUI()
                        chkCount = 0;
                        chkReadyState();
                    }
                    --%>

                    return 1;
                });

                // 設定 iframe 高度
                setMainFrameHeight();

                // 解除 menuClick() 及 form submit() 觸發的 blockUI()
                unblockUI();

                // 抓取 contentDoc 中的 Result 及 Error Message 進行顯示
                lastErrorMessage = contentDoc.find("#Msg_LastErrorMessage").html();
                lastResultMessage = contentDoc.find("#Msg_LastResultMessage").html();
                redirectUrl = contentDoc.find("#Msg_RedirectUrlAfterBlock").html();

                CheckPopResultMessage();
            });
        });

        function setMainFrameHeight() {
            if (contentDoc == undefined) {
                debugTrace("setMainFrameHeight: contentDoc undefined");
                return;
            }
            var iHeight = contentDoc.find("form").height();
            if (iHeight == undefined) { iHeight = 666; }
            //debugTrace("iHeight: " + iHeight);
            //debugTrace("contentDoc Height: " + iHeight);
            //$('#mainFrame').height(iHeight + 60);

            //20180809、20180810
            var contentWin = document.getElementById('mainFrame').contentWindow;
            var contentUrl = contentWin.location.href;
            var myIndex1 = contentUrl.indexOf("main2");
            var myIndex2 = contentUrl.indexOf("main2_detail");
            //debugger;
            //debugTrace("contentWin: " + contentWin);
            //debugTrace("contentUrl: " + contentUrl);
            //debugTrace("myIndex1: " + myIndex1);
            //debugTrace("myIndex2: " + myIndex2);
            if (myIndex1 > -1) {
                $('#mainFrame').height(iHeight - 7);       //(依照main2頁面,暫時調整的數值)
            }
            else {
                if (iHeight <= 0) { iHeight = 300; }
                if (myIndex2 > -1) {
                    $('#mainFrame').height(iHeight + 70);  //(依照main2_detail頁面,暫時調整的數值)
                }
                else {
                    $('#mainFrame').height(iHeight + 60);  //(ERROR-其餘頁面,暫時調整的數值)
                }
            }
            //debugTrace("#mainFrame Height: " + $('#mainFrame').height());
        }

        <%-- 
        匯出 excel (button class 包含 "asp_Export_M") 之類的 submit blockUI, 因不會有 contentDoc onload 事件, 用來解除 blockUI 的到數計時設定
        --%>
        const UNLCOK_TIME_OUT = 3; //秒  30*100毫秒=3000毫秒=3秒  
        var chkCount = 0;
        function chkReadyState() {
            //var contentWin = document.getElementById('mainFrame').contentWindow; //var contentDoc = contentWin.document;
            //debugTrace("chkReadyState(" + chkCount + "): " + contentWin.readyState + ":" + contentDoc.readyState);
            chkCount++;
            if (chkCount >= UNLCOK_TIME_OUT * 10 /*|| contentDoc.readyState == "complete"*/) {
                debugTrace("time's up, unblockUI");
                unblockUI();
            }
            else {
                setTimeout(function () { chkReadyState(); }, 100);
            }
        }
        /* //keycode 120 //F9
        function helpKeydown() {
            //if (event.keyCode == 120) { open_help_1(); }
        }
        $("#MasterForm").keypress(function () {
            //helpKeydown();
            //console.log("Handler for .keypress() called."); debugger;
        });
        $random1 = getRandom(9999999);
        window.open('./Doc/HELP/SYS_03_022_06.pdf', 'hf' + $random1.toString(), 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
        return false;
        function open_help_1() {
            $s_dh1 = './Doc/HELP/';
            $random1 = getRandom(9999999);
            $o_helppdf1 = $('#Hid_helppdf1');
            if (!$o_helppdf1) { return false; }
            $Hid_helppdf1 = $s_dh1 + $o_helppdf1.val() + "?r=" + $random1.toString();
            window.open($Hid_helppdf1, 'hf' + $random1.toString(), 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
            return false;
        }
        */
        function getRandom(x) {
            /*return (Math.floor(Math.random() * x) + 1);*/
            return (Math.floor((window.crypto.getRandomValues(new Uint32Array(1))[0] / 4294967296) * x) + 1);
        }
        function open_help_1() {
            $s_nopdf_msg = "該功能沒有線上說明文件!";
            // contentWin 中的 document 物件
            //debugTrace("setMainFrameHeight: contentDoc undefined");
            if (contentDoc == undefined) {
                //debugTrace("open_help_1: contentDoc == undefined");
                return;
            }

            var contentWin = document.getElementById('mainFrame').contentWindow;
            contentDoc = $(contentWin.document);
            var frameTable = contentDoc.find("#FrameTable");
            if (frameTable != undefined) {
                $s_dh1 = './Doc/HELP/';
                $random1 = getRandom(9999999);
                var s_helppdf1 = contentDoc.find("#Msg_helppdf1").html();
                if (s_helppdf1 == "") {
                    //debugTrace("open_help_1: s_helppdf1 == Empty"); alert($s_nopdf_msg);
                    blockAlert($s_nopdf_msg);
                    return false;
                }
                if (s_helppdf1.length == 0) {
                    //debugTrace("open_help_1: s_helppdf1.length == 0"); alert($s_nopdf_msg);
                    blockAlert($s_nopdf_msg);
                    return false;
                }
                //$o_helppdf1 = $('#Msg_helppdf1');
                //if (!$o_helppdf1) { return false; }
                //if (typeof ($o_helppdf1) == "undefined") { return false; }
                //if ($o_helppdf1.val() == "undefined") { return false; }
                //alert('helppdf1 :' + helppdf1); debugger;
                $s_helppdf1_x = $s_dh1 + s_helppdf1 + "?r=" + $random1.toString();
                //debugTrace("open_help_1: s_helppdf1_x=" + $s_helppdf1_x);
                window.open($s_helppdf1_x, 'hf' + $random1.toString(), 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
                return false;
            }
            //debugTrace("open_help_1: frameTable == undefined");
        }
        function sessionTimeout() {
            blockAlert(timeoutMsg, "逾時", function () {
                window.location.href = "<%=Request.ApplicationPath %>MOICA_login";
            });
        }
        function subMenuShow(a) {
            debugTrace("subMenuShow: " + $(a).text());
            hideAllSubMenu(a);
            var ul = $(a).parent().find("ul").first();
            //debugTrace(ul);
            ul.removeClass("hide");
        }
        function hideAllMenu() {
            $(".navigation ul").find("ul").addClass("hide");
        }
        function hideAllSubMenu(a) {
            $(a).closest("ul").find("ul").addClass("hide");
        }
    </script>
</asp:Content>
