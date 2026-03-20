<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MOICA_upload.aspx.vb" Inherits="WDAIIP.upload" %>

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
        //Cross-Frame Scripting ( 11294 )
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

            var lastErrorMessage = '<%=sm.LastErrorMessage%>';  //$("span#LastErrorMessage").html();
            var lastResultMessage = '<%=sm.LastResultMessage%>'; //$("span#LastResultMessage").html();
            var redirectUrl = '<%=sm.RedirectUrlAfterBlock%>'; //$("span#RedirectUrlAfterBlock").html();

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
    </script>
    <script type="text/javascript">
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
    </script>
</head>
<body class="bodybg">
    <form id="form1" class="form-horizontal" runat="server">
        <input type="hidden" name="csrf_token" value="<%= Session("CSRF_TOKEN") %>" />
        <%--83E8351EB57B03648DE038633242CFF796591E8DAB28E1441E93CEF40DA8572B8B7C20E0E18489B902B44E1BFC078278103F34A33F42BA46EEE23130CCEE257A601D99DBEABCEDA0D35F02EF9437DD238267A16EC964C250310881A57666CCC850C00A71BB57C1434EF907F3C23DDFE456E74D3D4322B79C724B1B47492A8820E4AD4CE16DFE97E2F76466D003F70D86F470B6A6CE24FB2E5B707E6DF45E101B134C7402915D7C5338D131968A647A454327C1E2BB5E23BEE684F634ED68CE9F247F8C0792F85CD18692CE6F7864C3D1--%>
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
                            <h3 class="loginTitleA"><img src="/images/icon-arrow.svg" alt="項目符號" />自然人憑證綁定驗證</h3>
                            <div class="col-sm-12">
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtUserId"><span class="mark-red">＊</span>帳號</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtUserId" runat="server" placeholder="請輸入您的帳號" MaxLength="20" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtUserPass"><span class="mark-red">＊</span>密碼</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtUserPass" runat="server" TextMode="Password" Columns="20" MaxLength="30" placeholder="請輸入您的密碼" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtVCode"><span class="mark-red">＊</span>圖型驗證碼 </label>
                                    <div class="col-sm-8">
                                        <asp:TextBox ID="txtVCode" runat="server" CssClass="form-control formbar-bg" placeholder="請輸入下方圖片中文字" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>

                                <div class="form-group">
                                    <div class="col-sm-8">
                                        <img id="vCode" src="/Common/ValidateCode?rand=<%=DateTime.Now.Ticks%>"
                                            alt="驗證碼圖片" class="loginVCode pull-right"
                                            data-toggle="tooltip" data-placement="top" title="產生新驗證碼" />
                                        <a href="/Common/ValidateCode?Audio=Y" target="frmPlayer" title="語音撥放驗證碼"><img id="playCode" alt="撥放圖示" src="/images/speaker.png" height="40" class="pull-right" /></a>


                                        <iframe name="frmPlayer" style="display: none;"></iframe>
                                    </div>
                                </div>

                                <div class="login-bottom-line">
                                    <asp:Button ID="bt_submit" runat="server" CssClass="btn btn-primary" Text="&nbsp;&nbsp;&nbsp;確認&nbsp;&nbsp;&nbsp;" />
                                    <button type="reset" class="btn btn-default">&nbsp;&nbsp;&nbsp;重設&nbsp;&nbsp;&nbsp;</button>
                                </div>

                            </div>
                        </div>
                    </div>

                </div>
                <!-- container end -->


                <!-- page body end -->
            </div>
        </div>

    </form>

    <span id="LastErrorMessage" style="display: none"><asp:Literal ID="Lit_LastErrorMessage" runat="server" /></span>
    <span id="LastResultMessage" style="display: none"><asp:Literal ID="Lit_LastResultMessage" runat="server" /></span>
    <span id="RedirectUrlAfterBlock" style="display: none"><asp:Literal ID="Lit_RedirectUrlAfterBlock" runat="server" /></span>
</body>
</html>
