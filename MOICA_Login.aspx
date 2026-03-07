<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MOICA_Login.aspx.vb" Inherits="WDAIIP.MOICA_Login" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="x-ua-compatible" content="IE=11" />
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>勞動部勞動力發展署｜在職訓練資訊管理系統</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="description" content="勞動部勞動力發展署｜在職訓練資訊管理系統" />
    <meta name="keywords" content="勞動部,勞動力發展署,在職訓練資訊管理系統,資訊系統,自辦在職,產業人才投資" />
    <meta name="author" content="東柏資訊" />
    <meta name="copyright" content="本網頁著作權屬勞動部勞動力發展署所有" />
    <link href="/Content/jquery-confirm.min.css" rel="stylesheet" />
    <link href="/Content/bootstrap3-3-6.min.css" rel="stylesheet" />
    <link href="/Content/bootstrap-treeview.css" rel="stylesheet" />
    <link href="/Content/font-awesome.min.css" rel="stylesheet" />
    <link href="/css/base.css" rel="stylesheet" />
    <script type="text/javascript" src="./Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="./Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="./Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="./Scripts/jquery.blockUI.js"></script>
    <script type="text/javascript" src="./Scripts/bootstrap.js"></script>
    <script type="text/javascript" src="./Scripts/bootstrap-treeview.js"></script>
    <script type="text/javascript" src="./Scripts/global.js"></script>
    <script type="text/javascript" src="./Scripts/HiPKIErrorcode.js"></script>
    <script type="text/javascript" src="./Scripts/HiPKICerts.js"></script>
    <style id="antiClickjack" type="text/css">
        html { display: none; }
        body { display: none !important; background-color: #ffffff; }
    </style>
    <style type="text/css">
        .mask-on-jobs { opacity: 0.24; position: absolute; top: 0; left: 0; width: 100%; height: 100%; background-position: center; background-repeat: no-repeat; background-size: cover; background-image: url('images/wda_logo_on_job_bg-2.svg'); }
        .mask-on-jobs-bg { opacity: 0.48; background-position: center; background-repeat: no-repeat; background-size: inherit; background-image: url('images/wda_logo_on_job_bg-1.svg'); }
        .mask-on-jobs-bg2 ~ * { position: relative; }
        .header { background-color: #fff; position: relative; }
    </style>
    <% 
        Server.ScriptTimeout = 10
    %>
    <script type="text/javascript">
        //解決不支援 X-Frame-Options設定，需額外判斷
        //if (top != self) { top.location = self.location; }
        if (self === top || self == top) {
            var antiClickjack = document.getElementById("antiClickjack");
            if (antiClickjack) { antiClickjack.parentNode.removeChild(antiClickjack); }
            document.documentElement.style.display = 'block';
        }
        else { top.location = self.location; }
        if (parent.document.frames != undefined && parent.document.frames.length != 0) {
            top.location.replace(self.location);
        }
        //alert(top.location);alert(self.location);

        //if (self == top) { document.documentElement.style.display = 'block'; }
        //else { top.location = self.location; }
    </script>
    <script type="text/javascript">
        //alert("openhttps:" + openhttps.value);
        //Hid_URLNG1/openhttps
        /*
        var Hid_URLNG1 = document.getElementById("Hid_URLNG1");
        if (Hid_URLNG1) {
            var urlNG1 = Hid_URLNG1.value;
            var urlw = window.location.href.toLowerCase(); //alert(url);
            var flag_empty = false;
            var flag_checkurl = true;
            if (urlNG1 == "") { flag_empty = true; }
            if (!flag_empty) { if (urlw.indexOf(urlNG1) != -1) { flag_checkurl = false; } }
            if (urlw.indexOf("https:") == -1 && urlw.indexOf("localhost") == -1 && flag_checkurl) {
                var Usehttps = false;
                var openhttps = document.getElementById("openhttps");
                if (openhttps) {
                    if (openhttps.value != "0") { Usehttps = true; }
                }
                else {
                    Usehttps = true;
                }
                if (Usehttps) {
                    url = url.replace("http:", "https:");
                    window.location.replace(url);
                }
            }
        }
        */

        $(document).ready(function () {
            var lastErrorMessage = $("span#LastErrorMessage").html();
            var lastResultMessage = $("span#LastResultMessage").html();
            var redirectUrl = $("span#RedirectUrlAfterBlock").html();
            if (window.top && window.top.sessionTimeout) {
                // 跑到這裡應該是 session time
                parent.sessionTimeout();
                return;
            }
            if (lastErrorMessage) {
                blockAlert(lastErrorMessage, "錯誤訊息");
            }
            else if (lastResultMessage) {
                blockMessage(lastResultMessage);
            }
        });

        function reloadValidCode() {
            blockUI();
            $('#vCode').attr("src", '/Common/ValidateCode' + "?rand=" + new Date().getMilliseconds());
        }

        $(function () {
            $('#vCode').on("click", function (e) {
                reloadValidCode();
            });
            $('#vCode').on("load", function () {
                unblockUI();
            });
            $('[data-toggle="tooltip"]').tooltip()
        });

        var certVerifyOk = 0;

        <%-- 執行輸入驗證 --%>
        function dosign() {
            var msg = "";
            var IDNo = $("input#txtIDNO").val();
            var PinCode = $("input#txtPin").val();
            if (!IDNo) {
                msg += "請輸入正確的個人身分證號(或居留證號)，不可為空。<br/>";
            }
            if (!PinCode) {
                msg += "請輸入正確的PIN 密碼，不可為空。<br/>";
            }
            IDNo = IDNo.toUpperCase();
            if (!checkId(IDNo) && !checkId2(IDNo) && !checkId4(IDNo)) {
                msg += "請輸入正確的個人身分證號(或居留證號)格式。<br/>";
            }

            //居留證(統一證)編號
            if (msg != "") {
                blockAlert(msg);
                return false;
            }

            $("input#txtIDNO").val(IDNo);
            if (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) {

                $("#waiting").show("fast", "swing", function () {
                    <%-- 呼叫自然人跨平台元件驗證Pin碼 --%>
                    CertValidation(PinCode, CertCallback);
                });
            }
            else {
                <%-- 呼叫自然人跨平台元件驗證Pin碼 --%>
                CertValidation(PinCode, CertCallback);
            }
        }

        <%-- 自然人跨平台元件驗證後會呼叫這個 function --%>
        function CertCallback(certData, rtnCode, msg) {
            $("#waiting").hide();
            if (rtnCode != 0) {
                blockAlert(msg, "憑證驗證失敗");
                return;
            }
            if (!certData) {
                blockAlert("跨平台元件 CertValidation() 失敗: 沒有回傳憑證資料!");
                return;
            }
            var name = "";
            var serial = "";
            var lastfour = certData.subjectID;
            var dnPairs = certData.subjectDN.split(",");
            for (var i = 0; i < dnPairs.length; i++) {
                var tokens = dnPairs[i].split("=");
                if (tokens.length == 2) {
                    if (console) console.log("tokens[0]:" + tokens[0] + ",tokens[1]:" + tokens[1]);
                    if (tokens[0] == "CN") {
                        name = tokens[1];
                    }
                    else if (tokens[0] == "serialNumber") {
                        serial = tokens[1];
                    }
                }
            }

            //var msg = 'name=' + name + '\nserialnumber=' + serial + '\nlast four=' + lastfour;
            var msg1 = name + '~~' + serial + '~~' + lastfour;
            var formObj = $("form#form1");
            formObj.find("#Hide_sign").val(certData.signedData);
            formObj.find("#Hide_enccert").val(certData.certB64);
            formObj.find("#Hide_cadata").val(msg1);
            formObj.find("#Hide_PinVerify").val("Ok");
            certVerifyOk = 1;
            if (console) console.log("Cert Verfify Success");
            formObj.submit();
        }

        /* 檢查輸入的身分證字號是否正確
		* @param   IDString	欲檢查的身分證字號
		* @return  boolean */
        function checkId(IDString) {
            var ID1 = (IDString ? IDString.toUpperCase() : "");
            if (ID1.length != 10) return false; //alert("身分證字號字數不對 !");
            var IDdigit = new Array(10);
            for (var i = 0; i < 10; i++) { IDdigit[i] = ID1.charAt(i); }
            var CharEng = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            IDdigit[0] = CharEng.indexOf(IDdigit[0]);
            if (IDdigit[0] == -1) return false; //alert("身分證字號第一位為錯誤英文字母 !");
            if (IDdigit[1] != 1 && IDdigit[1] != 2) return false; //alert("身分證字號無法辨識性別 !");
            var Array1 = new Array(26);
            Array1[0] = 1; Array1[1] = 10; Array1[2] = 19;
            Array1[3] = 28; Array1[4] = 37; Array1[5] = 46;
            Array1[6] = 55; Array1[7] = 64; Array1[8] = 39;
            Array1[9] = 73; Array1[10] = 82; Array1[11] = 2;
            Array1[12] = 11; Array1[13] = 20; Array1[14] = 48;
            Array1[15] = 29; Array1[16] = 38; Array1[17] = 47;
            Array1[18] = 56; Array1[19] = 65; Array1[20] = 74;
            Array1[21] = 83; Array1[22] = 21; Array1[23] = 3;
            Array1[24] = 12; Array1[25] = 30;
            var result = Array1[IDdigit[0]];
            for (var i = 1; i < 10; i++) {
                var Number = "0123456789";
                IDdigit[i] = Number.indexOf(IDdigit[i]);
                if (IDdigit[i] == -1) {
                    //alert("身分證字號錯誤 !");
                    return false;
                } else {
                    result += IDdigit[i] * (9 - i);
                }
            }
            result += 1 * IDdigit[9];
            //alert("result=="+result);
            //alert("身分證字號錯誤 !");
            if (result % 10 != 0) return false;
            return true;
        }

        /*
        外籍生，居留證號規則跟身分證號差不多，只是第二碼也是英文字母代表性別，跟第一碼轉換二位數字規則相同，但只取餘數
        */
        function checkId2(studIdNumber) {
            //外籍生，居留證號規則跟身分證號差不多，只是第二碼也是英文字母代表性別，跟第一碼轉換二位數字規則相同，但只取餘數
            //驗證填入身分證字號長度及格式 //alert("長度不足");
            if (studIdNumber.length != 10) return false;
            //格式，用正則表示式比對第一個字母是否為英文字母 //alert("格式錯誤");
            if (isNaN(studIdNumber.substr(2, 8)) || (!/^[A-Z]$/.test(studIdNumber.substr(0, 1))) || (!/^[A-Z]$/.test(studIdNumber.substr(1, 1)))) {
                return false;
            }

            var idHeader = "ABCDEFGHJKLMNPQRSTUVXYWZIO"; //按照轉換後權數的大小進行排序
            //這邊把身分證字號轉換成準備要對應的
            studIdNumber = (idHeader.indexOf(studIdNumber.substring(0, 1)) + 10)
                + '' + ((idHeader.indexOf(studIdNumber.substr(1, 1)) + 10) % 10)
                + '' + studIdNumber.substr(2, 8);
            //開始進行身分證數字的相乘與累加，依照順序乘上1987654321

            s = parseInt(studIdNumber.substr(0, 1)) +
                parseInt(studIdNumber.substr(1, 1)) * 9 +
                parseInt(studIdNumber.substr(2, 1)) * 8 +
                parseInt(studIdNumber.substr(3, 1)) * 7 +
                parseInt(studIdNumber.substr(4, 1)) * 6 +
                parseInt(studIdNumber.substr(5, 1)) * 5 +
                parseInt(studIdNumber.substr(6, 1)) * 4 +
                parseInt(studIdNumber.substr(7, 1)) * 3 +
                parseInt(studIdNumber.substr(8, 1)) * 2 +
                parseInt(studIdNumber.substr(9, 1));

            //檢查號碼 = 10 - 相乘後個位數相加總和之尾數。
            checkNum = parseInt(studIdNumber.substr(10, 1));
            //模數 - 總和/模數(10)之餘數若等於第九碼的檢查碼，則驗證成功
            ///若餘數為0，檢查碼就是0
            if ((s % 10) == 0 || (10 - s % 10) == checkNum) return true;
            return false;
        }

        //'4:居留證2(外來人口統一證號)／新式統一證號 
        function checkId4(studIdNumber) {
            if (studIdNumber.length != 10) return false;
            //9碼為數字 //第1個字母是否為英文字母
            if (isNaN(studIdNumber.substr(1, 9)) || !/^[A-Z]$/.test(studIdNumber.substr(0, 1))) return false;
            //按照轉換後權數的大小進行排序
            var idHeader = "ABCDEFGHJKLMNPQRSTUVXYWZIO"; 
            //這邊把身分證字號轉換成準備要對應的
            studIdNumber = (idHeader.indexOf(studIdNumber.substring(0, 1)) + 10)
                + '' + studIdNumber.substr(1, 9);
            //開始進行身分證數字的相乘與累加，依照順序乘上1987654321
            var intS = parseInt(studIdNumber.substr(0, 1))
                + parseInt(studIdNumber.substr(1, 1)) * 9
                + parseInt(studIdNumber.substr(2, 1)) * 8
                + parseInt(studIdNumber.substr(3, 1)) * 7
                + parseInt(studIdNumber.substr(4, 1)) * 6
                + parseInt(studIdNumber.substr(5, 1)) * 5
                + parseInt(studIdNumber.substr(6, 1)) * 4
                + parseInt(studIdNumber.substr(7, 1)) * 3
                + parseInt(studIdNumber.substr(8, 1)) * 2
                + parseInt(studIdNumber.substr(9, 1)) * 1
                + parseInt(studIdNumber.substr(10, 1)); //(檢查碼)
            //證號OK
            if ((intS % 10) == 0) return true;
            return false;
        }

        function transferLogin() {
            document.location.href = "login";
        }

        function winPopUp1() {
            //var HiPKULocalServer = "http://localhost:61161";
            var windowName = 'userConsole';
            var popUp = window.open("js/OpenWin/wPopUp1.aspx", windowName, "height=20,width=66,left=33,top=33,scrollbars,resizable");
            if (popUp == null || typeof (popUp) == 'undefined') {
                //瀏覽器設定須允許本系統的 PopUp 視窗
                blockAlert('請解除視窗阻攔(須允許本系統的 PopUp 視窗)，重新點選連結。', "錯誤訊息");
            }
            else {
                popUp.focus();
            }
        }
    </script>
</head>
<body class="bodybg">
    <form id="form1" class="form-horizontal" method="post" runat="server">
        <%--<span id="httpWrapper" runat="server"></span>--%>
        <%--<input id="openhttps" size="1" type="hidden" name="openhttps" runat="server" autocomplete="off" />--%>
        <asp:HiddenField ID="Hide_PinVerify" runat="server" />
        <%--<asp:HiddenField ID="Hide_vcode" runat="server" />--%>
        <asp:HiddenField ID="Hide_enccert" runat="server" />
        <asp:HiddenField ID="Hide_cadata" runat="server" />
        <asp:HiddenField ID="Hide_sign" runat="server" />
        <%--<input id="nonce" type="hidden" runat="server" name="nonce" autocomplete="off" />--%>
        <div class="bodybg">
            <div class="container-fluid">
                <!-- page body start -->
                <div class="container">
                    <!-- header start -->
                    <div class="header" style="height: 80px; text-align: center;">
                        <div class="mask-on-jobs"></div>
                        <div class="mask-on-jobs-bg"></div>
                        <div class="logo-login">
                            <img src="images/wda_logo_on_job.svg" class="img-responsive" alt="勞動部勞動力發展署職業訓練資訊管理系統" />
                        </div>
                    </div>
                    <!-- header end -->
                    <div class="col-sm-6 col-md-offset-3">
                        <div class="login-bar">
                            <h3 class="loginTitleA">
                                <img src="/images/icon-arrow.svg" alt="項目符號" />自然人憑證登入</h3>
                            <div class="col-sm-12">
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtIDNO"><span class="mark-red">＊</span>身分證字號</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtIDNO" runat="server" TextMode="Password" placeholder="請輸入您的身分證字號(或居留證號)" MaxLength="20" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtPin"><span class="mark-red">＊</span>PIN 碼</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtPin" runat="server" TextMode="Password" Columns="20" MaxLength="30" placeholder="請輸入自然人憑證 PIN 碼" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <%-- 
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtVCode"><span class="mark-red">＊</span>圖型驗證碼</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox ID="txtVCode" runat="server" CssClass="form-control formbar-bg" placeholder="請輸入下方圖片中文字" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <div class="col-sm-8">
                                        <img id="vCode" src="/Common/ValidateCode?<%=DateTime.Now.Ticks%>"
                                                alt="驗證碼圖片" class="loginVCode pull-right"
                                                data-toggle="tooltip" data-placement="top" title="產生新驗證碼" />
                                        <a href="/Common/ValidateCode?Audio=Y" target="frmPlayer" title="語音撥放驗證碼"><img id="playCode" alt="撥放圖示" src="/images/speaker.png" height="40" class="pull-right" /></a>
                                        <iframe name="frmPlayer" style="display:none;"></iframe>
                                    </div>
                                </div>
                                --%>
                                <div class="login-bottom-line">
                                    <button type="button" class="btn btn-primary" onclick="dosign()">&nbsp;&nbsp;&nbsp;登入&nbsp;&nbsp;&nbsp;</button>
                                    <button type="reset" class="btn btn-default">&nbsp;&nbsp;&nbsp;重設&nbsp;&nbsp;&nbsp;</button>
                                    <button type="button" class="btn btn-default" onclick="transferLogin()" id="btnPWDLOGIN" runat="server" visible="False">&nbsp;&nbsp;&nbsp;帳號密碼登入&nbsp;&nbsp;&nbsp;</button>
                                </div>
                            </div>
                            <div>
                                注意事項：
                    <ol>
                        <li>請確認您已安裝成功最新版本「<a href='<%= WDAIIP.TIMS.Get_downloadMain%>' target="_blank">內政部憑證管理中心-跨平台網頁元件</a>」。</li>
                        <li>請確認「跨平台網頁元件服務」已正確啟動。</li>
                        <li>請確認已正確連接讀卡機並插入自然人憑證。<a target="_blank" href="http://localhost:61161/selfTest.htm">test</a></li>
                        <li>瀏覽器設定須允許本系統的 <a id="APopUp1" title="PopUp 視窗" href="#" onclick="winPopUp1();">PopUp 視窗</a>。</li>
                        <li>若您連續三次輸入錯誤PIN碼，將會造成鎖卡喔!!</li>
                        <li>內政部官方解除卡片鎖卡網頁：<br />
                            <a target="_blank" href="<%= WDAIIP.TIMS.Get_unblockcard%>"><%= WDAIIP.TIMS.Get_unblockcard%></a>
                        </li>
                    </ol>
                            </div>
                        </div>
                    </div>
                </div>
                <!-- container end -->
                <!-- page body end -->
            </div>
        </div>
        <div style="position: absolute; top: 508px; left: 1px;" id="div12" runat="server">
            <%--ForeColor="White"--%>
            <asp:Label ID="Labmsg1" runat="server" Text="Labx" ForeColor="#D5EEFF"></asp:Label>
        </div>
        <div style="position: absolute; left: -22px; top: 12px; height: 17px; width: 36px;" id="divC" runat="server">
            <a id="A1" title="關閉" class="l" href="#" onclick="window.opener=null; window.open('','_self'); window.close();" style="color: #D5EEFF">關閉</a>
        </div>
        <%--<asp:HiddenField ID="Hid_URLNG1" runat="server" />--%>
    </form>
    <%--
    <span id="LastErrorMessage" style="display: none"><asp:Literal ID="Lit_LastErrorMessage" runat="server" /></span>
    <span id="LastResultMessage" style="display: none"><asp:Literal ID="Lit_LastResultMessage" runat="server" /></span>
    <span id="RedirectUrlAfterBlock" style="display: none"><asp:Literal ID="Lit_RedirectUrlAfterBlock" runat="server" /></span>
    <div id="waiting" style="display: none; width: 200px; height: 200px; left: 100px; top: 20px; position: absolute; z-index: 999; background-color: #FFF; border: 2px solid #808080; text-align: center; padding-top: 15px">
        <span>憑證檢核中</span><br />
        <img src="/images/waiting.gif" alt="憑證檢核中" />
    </div>
    --%>
    <span id="LastErrorMessage" style="display: none"><asp:Literal ID="Lit_LastErrorMessage" runat="server" /></span>
    <span id="LastResultMessage" style="display: none"><asp:Literal ID="Lit_LastResultMessage" runat="server" /></span>
    <span id="RedirectUrlAfterBlock" style="display: none"><asp:Literal ID="Lit_RedirectUrlAfterBlock" runat="server" /></span>
    <div id="waiting" style="display: none; width: 200px; height: 200px; left: 100px; top: 20px; position: absolute; z-index: 999; background-color: #FFF; border: 2px solid #808080; text-align: center; padding-top: 15px">
        <span>憑證檢核中</span><br />
        <img src="/images/waiting.gif" alt="憑證檢核中" />
    </div>
</body>
</html>
