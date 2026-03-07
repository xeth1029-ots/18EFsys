<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Login.aspx.vb" Inherits="WDAIIP.Login" %>

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
    <link href="~/Content/jquery-confirm.min.css" rel="stylesheet" />
    <link href="~/Content/bootstrap3-3-6.min.css" rel="stylesheet" />
    <link href="~/Content/bootstrap-treeview.css" rel="stylesheet" />
    <link href="~/Content/font-awesome.min.css" rel="stylesheet" />
    <link href="~/css/base.css" rel="stylesheet" />
    <script type="text/javascript" src="<%=BaseUrl%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=BaseUrl%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="<%=BaseUrl%>Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="<%=BaseUrl%>Scripts/jquery.blockUI.js"></script>
    <script type="text/javascript" src="<%=BaseUrl%>Scripts/bootstrap.js"></script>
    <script type="text/javascript" src="<%=BaseUrl%>Scripts/bootstrap-treeview.js"></script>
    <script type="text/javascript" src="<%=BaseUrl%>Scripts/global.js"></script>
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

        $(document).ready(function () {
            var isLogin = '<% If sm.IsLogin Then Response.Write("1") Else Response.Write("0") %>';
            var lastErrorMessage = '<%=sm.LastErrorMessage%>';
            var lastResultMessage = '<%=sm.LastResultMessage%>';
            var redirectUrl = '<%=sm.RedirectUrlAfterBlock%>';
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

        function transferLogin() {
            document.location.href = "MOICA_Login";
        }
    </script>
    <script type="text/javascript">
        function reloadValidCode() {
            blockUI();
            $('#vCode').attr("src", '<%=BaseUrl%>Common/ValidateCode' + "?rand=" + new Date().getMilliseconds());
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
    </script>
</head>
<body class="bodybg">
    <form id="form1" class="form-horizontal" runat="server">
        <%--<input id="openhttps" size="1" type="hidden" name="openhttps" runat="server" autocomplete="off" />--%>
        <div class="bodybg">
            <div class="container-fluid">
                <!-- page body start -->
                <div class="container">
                    <!-- header start -->
                    <div class="header" style="height: 80px; text-align: center;">
                        <div class="mask-on-jobs"></div>
                        <div class="mask-on-jobs-bg"></div>
                        <div class="logo-login">
                            <img src="<%=BaseUrl%>images/wda_logo_on_job.svg" class="img-responsive" alt="勞動部勞動力發展署職業訓練資訊管理系統" />
                        </div>
                    </div>
                    <!-- header end -->
                    <div class="col-sm-6 col-md-offset-3">
                        <div class="login-bar">
                            <h3 class="loginTitleA">
                                <img src="<%=BaseUrl%>images/icon-arrow.svg" alt="項目符號" />系統登入</h3>
                            <div class="col-sm-12">
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtUserId"><span class="mark-red2">＊</span>帳號</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtUserId" runat="server" placeholder="請輸入您的帳號" MaxLength="20" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtUserPass"><span class="mark-red2">＊</span>密碼</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtUserPass" runat="server" TextMode="Password" Columns="20" MaxLength="30" placeholder="請輸入您的密碼" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtVCode"><span class="mark-red2">＊</span>圖型驗證碼 </label>
                                    <div class="col-sm-8">
                                        <asp:TextBox ID="txtVCode" runat="server" CssClass="form-control formbar-bg" placeholder="請輸入下方圖片中文字" AutoComplete="off" MaxLength="10"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <div class="col-sm-8">
                                        <img id="vCode" src="<%=BaseUrl%>Common/ValidateCode?rand=<%=DateTime.Now.Ticks%>" alt="驗證碼圖片" class="loginVCode pull-right" data-toggle="tooltip" data-placement="top" title="產生新驗證碼" />
                                        <a href="/Common/ValidateCode?Audio=Y" target="frmPlayer" title="語音撥放驗證碼">
                                            <img id="playCode" alt="撥放圖示" src="<%=BaseUrl%>images/speaker.png" height="40" class="pull-right" /></a>
                                        <iframe name="frmPlayer" style="display: none;"></iframe>
                                    </div>
                                </div>
                                <div class="login-bottom-line">
                                    <asp:Button ID="bt_submit" runat="server" CssClass="btn btn-primary" Text="&nbsp;&nbsp;&nbsp;登入&nbsp;&nbsp;&nbsp;" />
                                    <asp:Button ID="bt_reset" type="reset" runat="server" CssClass="btn btn-default" Text="&nbsp;&nbsp;&nbsp;重設&nbsp;&nbsp;&nbsp;" />
                                    <asp:Button ID="bt_FRGTPXSWXD" type="reset" runat="server" CssClass="btn btn-default" Text="&nbsp;&nbsp;&nbsp;忘記密碼&nbsp;&nbsp;&nbsp;" />
                                    <%--<button type="reset" class="btn btn-default">&nbsp;&nbsp;&nbsp;重設&nbsp;&nbsp;&nbsp;</button>--%>
                                    <%--<asp:Button ID="bt_MOICA" runat="server" CssClass="btn btn-default" Text="&nbsp;&nbsp;&nbsp;自然人憑證登入&nbsp;&nbsp;&nbsp;" />--%>
                                    <button type="button" class="btn btn-default" onclick="transferLogin()">&nbsp;&nbsp;&nbsp;自然人憑證登入&nbsp;&nbsp;&nbsp;</button>
                                    <asp:Button ID="bt_atest" runat="server" CssClass="btn btn-default" Text="&nbsp;&nbsp;&nbsp;ATEST&nbsp;&nbsp;&nbsp;" />
                                    <asp:HiddenField ID="Hidversion1" runat="server" />
                                </div>
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
        <div style="position: absolute; left: -20px; top: 22px; height: 17px; width: 36px;" id="divC" runat="server">
            <a id="A1" title="關閉" class="l" href="#" onclick="window.opener=null; window.open('','_self'); window.close();" style="color: #D5EEFF">關閉</a>
        </div>
        <%--<asp:HiddenField ID="Hid_URLNG1" runat="server" />--%>
    </form>
</body>
</html>
