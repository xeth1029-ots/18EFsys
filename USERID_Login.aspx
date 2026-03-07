<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="USERID_Login.aspx.vb" Inherits="WDAIIP.USERID_Login" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>勞動部勞動力發展署｜產業人材投資方案資訊管理系統</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="description" content="勞動部勞動力發展署｜產業人材投資方案資訊管理系統" />
    <meta name="keywords" content="勞動部,勞動力發展署,產業人材投資方案資訊管理系統,資訊系統,自辦在職,產業人材投資" />
    <meta name="author" content="東柏資訊" />
    <meta name="copyright" content="本網頁著作權屬勞動部勞動力發展署所有" />

    <link href="~/Content/jquery-confirm.min.css" rel="stylesheet"/>
    <link href="~/Content/bootstrap3-3-6.min.css" rel="stylesheet"/>
    <link href="~/Content/bootstrap-treeview.css" rel="stylesheet"/>
    <link href="~/Content/font-awesome-4.7.0.min.css" rel="stylesheet"/>

    <link href="~/css/base.css" rel="stylesheet"/>

    <script type="text/javascript" src="~/Scripts/jquery-1.10.2.js"></script>
    <script type="text/javascript" src="~/Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="~/Scripts/jquery.blockUI.js"></script>

    <script type="text/javascript" src="~/Scripts/bootstrap.js"></script>
    <script type="text/javascript" src="~/Scripts/bootstrap-treeview.js"></script>

    <script type="text/javascript" src="/Scripts/global.js"></script>

    <script type="text/javascript">
        $(document).ready(function () {

            
            var lastErrorMessage = '';
            var lastResultMessage = '';
            var redirectUrl = '';
            if (lastErrorMessage) {
                blockAlert(lastErrorMessage, "錯誤訊息");
            }
            else if (lastResultMessage) {
                blockMessage(lastResultMessage, null, function () {
                    if (redirectUrl) {
                        
                        var form = $("form[name=RedirAfterBlockResult]");
                        var p = redirectUrl.indexOf("?");
                        if (p > 0) {
                            var parms = redirectUrl.substring(p + 1).split("&");
                            redirectUrl = redirectUrl.substring(0, p);
                            for (var i = 0; i < parms.length; i++) {
                                var ptoken = parms[i].split("=");
                                if (ptoken.length == 2) {
                                    var input = document.createElement("input");
                                    $(input).attr("type", "hidden");
                                    $(input).attr("name", ptoken[0]);
                                    $(input).attr("value", ptoken[1]);

                                    form.append(input);
                                }
                                else {
                                    debugTrace("RedirectUrlAfterBlock: parameter syntax error '" + parms[i] + "'" );
                                }
                            }
                        }
                        form.attr("action", redirectUrl);
                        form.submit();
                    }
                });
            }
            else {
                
                var IsValid = 'True';
                var validationMsg = $("#ValidationSummary").html();
                if (IsValid == 'False' && validationMsg) {
                    blockAlert(validationMsg, "表單檢核訊息");
                }
            }

        });
    </script>

	<script type="text/javascript">
		
		function btsubmitclick() {
			document.getElementById('bt_submit').click();
		}
		function RefreshImage(valImageId) {
			var objImage = document.getElementById(valImageId)
			if (objImage == undefined) {
				return;
			}
			var now = new Date();
			//alert(objImage.src.split('?')[0] + '?x=' + now.toUTCString());
			objImage.src = objImage.src.split('?')[0] + '?x=' + now.toUTCString();
		}
	</script>
</head>
<body>
    <form class="form-horizontal" method="post" runat="server">

    <div class="bodybg">
        <div class="container-fluid">
            <!-- page body start -->
            

<div class="container">
    <!-- header start -->
    <div class="header" style="background-image:url('/images/logo-index-bg.png');height:80px;">
        
        <center><div class="logo-login"><img src="/images/logo-index.png" class="img-responsive" alt="勞動部勞動力發展署職業訓練資訊管理系統" /></div></center>
    </div>
    <!-- header end -->
    <div class="col-sm-6 col-md-offset-3">
        <div class="login-bar">
            <h3 class="loginTitleA"><img src="/images/icon-arrow.svg" alt="項目符號" />系統登入</h3>
                <div class="col-sm-12">
                    <div class="form-group">
                        <label class="control-label col-sm-4 label-set" for="txtname"><span class="mark-red">＊</span>帳號</label>
                        <div class="col-sm-8">
                            <asp:TextBox CssClass="form-control formbar-bg" ID="txtname" runat="server" placeholder="請輸入您的帳號" MaxLength="20" AutoCompleteType="Disabled"></asp:TextBox>
                        </div>
                    </div>
                    <div class="form-group">
                        <label class="control-label col-sm-4 label-set" for="txtpass"><span class="mark-red">＊</span>密碼</label>
                        <div class="col-sm-8">
                            <asp:TextBox CssClass="form-control formbar-bg" ID="txtpass" runat="server" TextMode="Password" Columns="20" MaxLength="30" placeholder="請輸入您的密碼" AutoCompleteType="Disabled"></asp:TextBox>
                        </div>
                    </div>
                    <div class="form-group">
                        <label class="control-label col-sm-4 label-set" for="txtvnum"><span class="mark-red">＊</span>圖型驗證碼 </label>
                        <div class="col-sm-8">
                            <asp:TextBox ID="txtvnum" runat="server" placeholder="請輸入下方圖片中文字" AutoCompleteType="Disabled"></asp:TextBox>
                        </div>
                    </div>

                    <div class="form-group">
                        <div class="col-sm-8">
                            <img id="vCode" src="/Common/ValidateCode.aspx"
                                    alt="驗證碼" class="loginVCode pull-right"
                                    data-toggle="tooltip" data-placement="top" title="產生新驗證碼" />
                            <a href="VCodeAudio" target="frmPlayer" title="語音撥放驗證碼"><img id="playCode" alt="撥放" src="/images/speaker.png" height="40" class="pull-right" /></a>

                            
                            <iframe name="frmPlayer" style="display:none;"></iframe>
                        </div>
                    </div>

                    <div class="login-bottom-line">
                        <button runat="server" class="btn btn-primary" onserverclick="bt_submit_Click">&nbsp;&nbsp;&nbsp;登入&nbsp;&nbsp;&nbsp;</button>
                        <button type="button" class="btn btn-default" onclick="location.href = '';">
                            忘記密碼
                        </button>
                        <button type="button" class="btn btn-default" onclick="location.href = '';">
                            首次登入
                        </button>
                    </div>

                </div>
        </div>
    </div>
</div>
 


            <!-- page body end -->
        </div>
    </div>

        </form>
</body>
</html>
