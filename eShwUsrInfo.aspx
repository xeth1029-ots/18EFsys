<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="eShwUsrInfo.aspx.vb" Inherits="WDAIIP.eShwUsrInfo" %>

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
        <%--<input id="openhttps" size="1" type="hidden" name="openhttps" runat="server" autocomplete="off" />--%>
        <div class="bodybg">
            <div class="container-fluid">
                <!-- page body start -->
                <div class="container">
                    <!-- header start -->
                    <div class="header" style="background-image: url('/images/logo-index-bg.png'); height: 80px; text-align: center;">
                        <div class="logo-login">
                            <img src="/images/logo-index.png" class="img-responsive" alt="勞動部勞動力發展署職業訓練資訊管理系統" />
                        </div>
                    </div>
                    <!-- header end -->
                    <div class="col-sm-6 col-md-offset-3">
                        <div class="login-bar">
                            <h3 class="loginTitleA">
                                <img src="/images/icon-arrow.svg" alt="項目符號" />帳號資訊</h3>
                            <div class="col-sm-12">
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtOrgName">訓練機構</label>
                                    <div class="col-sm-8 white">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtOrgName" runat="server" MaxLength="20" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtUserName">姓名</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtUserName" runat="server" placeholder="請輸入您的姓名" MaxLength="20" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtUserId">帳號</label>
                                    <div class="col-sm-8 white">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtUserId" runat="server" MaxLength="20" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtUserEMAIL"><span class="mark-red">＊</span>E-Mail</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtUserEMAIL" runat="server" placeholder="請輸入您的E-Mail" MaxLength="60" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtUserPhone">電話</label>
                                    <div class="col-sm-8">
                                        <asp:TextBox CssClass="form-control formbar-bg" ID="txtUserPhone" runat="server" placeholder="請輸入您的電話" MaxLength="60" AutoComplete="off"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <label class="control-label col-sm-4 label-set" for="txtVCode"><span class="mark-red">＊</span>圖型驗證碼 </label>
                                    <div class="col-sm-8">
                                        <asp:TextBox ID="txtVCode" runat="server" CssClass="form-control formbar-bg" placeholder="請輸入下方圖片中文字" AutoComplete="off" MaxLength="10"></asp:TextBox>
                                    </div>
                                </div>
                                <div class="form-group">
                                    <div class="col-sm-8">
                                        <img id="vCode" src="/Common/ValidateCode?rand=<%=DateTime.Now.Ticks%>" alt="驗證碼圖片" class="loginVCode pull-right" data-toggle="tooltip" data-placement="top" title="產生新驗證碼" />
                                        <a href="/Common/ValidateCode?Audio=Y" target="frmPlayer" title="語音撥放驗證碼">
                                            <img id="playCode" alt="撥放圖示" src="/images/speaker.png" height="40" class="pull-right" /></a>
                                        <iframe name="frmPlayer" style="display: none;"></iframe>
                                    </div>
                                </div>
                                <div class="login-bottom-line">
                                    <asp:Button ID="bt_save1" type="save" runat="server" CssClass="btn btn-default" Text="&nbsp;&nbsp;&nbsp;儲存&nbsp;&nbsp;&nbsp;" />
                                    &nbsp;&nbsp;
                                    <asp:Button ID="bt_submit" runat="server" CssClass="btn btn-primary" Text="&nbsp;&nbsp;&nbsp;修改密碼&nbsp;&nbsp;&nbsp;" />
                                    &nbsp;
                                    <asp:Button ID="bt_back1" type="back1" runat="server" CssClass="btn btn-default" Text="&nbsp;&nbsp;&nbsp;回上頁&nbsp;&nbsp;&nbsp;" />

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
        <asp:HiddenField ID="Hid_acount_1" runat="server" />
        <%--<asp:HiddenField ID="Hid_URLNG1" runat="server" />--%>
    </form>
</body>
</html>
