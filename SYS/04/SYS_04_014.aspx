<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_014.aspx.vb" Inherits="WDAIIP.SYS_04_014" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>機構地址WGS84定位工具</title>
    <%--<script src="<%=ResolveUrl("~/Internet/index/js/jquery1.11.1.js")%>" type="text/javascript"></script>--%>
    <script type="text/javascript" src="<%=ResolveUrl("~/Scripts/jquery-3.7.1.min.js")%>"></script>
    <script type="text/javascript" src="<%=ResolveUrl("~/Scripts/jquery-migrate-3.4.1.min.js")%>"></script>
    <%--<script type="text/javascript" src="https://api.tgos.tw/TGOS_API/tgos?ver=2&AppID=rcaCK4WPtyOQheX1QlrBd33iWG45RnJyriVwNxb6uVQ3N0QrtVKWvw==&APIKey=cGEErDNy5yN/1fQ0vyTOZrghjE+jIU6uN8EzeAV7Xx16CxoUTagZUYxUYZXONCKjbpUgSS0mtLNbsYrfF+lZj6wNguZwA9F5kkwhBN/ONocxT7sAk6CvU2RcWiT7t/aYs4ycKA308hyhNyETwHOlmuuPYVIqXDjXOlalu6cnFHG/F9JQr2xVeVb501nrKgpQvsoVmttQi5IIE3laAjxlCXme66IZgcRjml65FaS7eh60mIIaSCfWNAEY2SJ3lKJyCIRfr/H7sAZQNClc3ijBVZHTkZ3tG/q3Zzq84LLSLTlsHRyV3tWP7zkWHiMsNSS0ci7ERWuGotvLvw7YLz/BGrFgitT9AGJsZuQx+A8/0iZiWeZI2jK6lRwxmH1ZTbQQSQCHaMrToKKj6d1URBKsAc1dlZ/FNE8FTWMJGHeM8b8=" charset="utf-8"></script>--%>
    <%--<script type="text/javascript" src="https://api.tgos.tw/TGOS_API/tgos?ver=2&AppID=lTPdox5XfzyjNW8l+oaklUZjW9jljaV3Lx3DOnYJKF37byXdJlfOmg==&APIKey=cGEErDNy5yN/1fQ0vyTOZrghjE+jIU6uuyLUpFfm0OEY+iBNJnf9WvoCwTopOOUPMfdDpg/pihiPALA5s/gl4J/KJZyWVApuKKNoJiAu3QlYzSJuGpzz9e6C1+N6l+HGG05E0LkMpa1VOeiBrb/BJbShahGt6WGkVFErM/NJ+eg2tIbOdFZFz8rNPtZJ5mGUo2vgXchvQ1mVNurt/Qx9/dtq0F1x66xmG6KpaFP/XNA5WXLSSy6dsjNneF4notc6ttmVaJiyiRJXWejNYPv/7UJ3QwcVfgMTWQ+Me6AtKZlgTlL7gVrd0YiSY6IVhvwrNCQ5L9RJ2xRyd8BPcMcW1oaZRl8x/ejt/AJJ0LI/L9AfXWci2CP7SSpzDfQPotpreF997vUiqCb1VrSYv4HrFDCMTupMZphHoXCdp9mmJIpuTioW70IGWg==" charset="utf-8"></script>--%>
    <script type="text/javascript" src="<%=WDAIIP.TIMS.Get_UrlStr("TGOS1") %>" charset="utf-8"></script>
    <script type="text/javascript" src="<%=ResolveUrl("~/js/tgos_map.js")%>" charset="utf-8"></script>
    <script type="text/javascript">
        var user_ip = '<%=Turbo.Common.GetIpAddress() %>';
        var locateOnGoing = false;
        var locateProcess = 0;
        var saveProcess = 0;
        var saveOnGoing = false;
        var total = 0;
        var idx = 0, i = 0;
        var addrList = [], data_list = [];
        var success = 0, failed = 0;

        var saveTotal = 0;
        var saveSuccess = 0, saveFailed = 0;

        $(document).ready(function () {

            $("ul input[type=checkbox]").on('click', function () {
                $(this).closest('ul').data('checked', $(this).is(":checked") ? "Y" : "N");
            });

            $("#btnCheckAll").on('click', function () {
                $("ul input[type=checkbox]").prop("checked", true);
                $("ul").data('checked', 'Y');
            });

            $("#btnUncheckAll").on('click', function () {
                $("ul input[type=checkbox]").prop("checked", false);
                $("ul").data('checked', 'N');
            });

            $("#btnLocate").on('click', function () {
                if (!locateOnGoing) { geoLocator(); }
            });

            $("#btnSave").on('click', function () {
                if (!saveOnGoing) { saveResults(); }
            })
        });

        function saveResults() {

            saveOnGoing = true;

            var jobULs = $("ul[id^=lv_classQueryResult]");
            var areanet = $("#txt_localareanet").val();

            saveTotal = 0;
            saveIdx = 0;

            // 計算待更新總筆數
            data_list = [];
            jobULs.each(function () {

                if ($(this).data("checked") != 'Y') { return; }

                var data = {};
                data.save = 'Y';
                data.ocid = $(this).data("id");
                data.wgs84_y = $(this).find("span[id=WGS84]").find("[id=Y]").val();
                data.wgs84_x = $(this).find("span[id=WGS84]").find("[id=X]").val();
                data.areanet = areanet;

                if (data.wgs84_y == "" || data.wgs84_x == "") { return; }

                data_list.push(data);
            });

            saveTotal = data_list.length;

            refreshSaveProgress();
            idx = 0;
            checkRunSave();
        }

        function checkRunSave() {
            while (saveProcess < 5 && idx < data_list.length) {
                saveProcess++;
                doSave(data_list[idx]);
                idx++;
            };

            if (idx < data_list.length) {
                setTimeout(function () {
                    checkRunSave();
                }, 100);
            }
        }

        function doSave(data) {
            //"../../ajax/SaveClassGeoLocator.ashx"
            var PostURL_1 = "<%=ResolveUrl("~/ajax/SaveClassGeoLocator.ashx")%>";
            var jqxhr = $.post(PostURL_1, data,
                    function (r) {
                        if (r.status == "ok") {
                            saveSuccess++;
                        }
                        else {
                            saveFailed++;
                            console.error("save failed: " + r.status + ", " + JSON.stringify(data));
                        }
                    })
                    .fail(function () {
                        saveFailed++;
                        console.error("save failed: " + JSON.stringify(data));
                    })
                    .always(function () {
                        saveProcess--;
                        refreshSaveProgress();
                    });
        }


        function refreshSaveProgress() {
            var progress = $("#litSaveProgress");
            if ((saveSuccess + saveFailed) < saveTotal) {
                progress.html("資料更新中 ... " + saveSuccess + " / " + saveFailed + " / " + saveTotal);
            }
            else {
                progress.html("資料更新作業完成, 共 " + saveTotal + " 筆 (成功 " + saveSuccess + " 筆, 失敗 " + saveFailed + " 筆)");
                saveOnGoing = false;
            }
        }

        function refreshProgress() {
            var progress = $("#litProgress");
            if ((success + failed) < total) {
                progress.html("地址定位中 ... " + success + " / " + failed + " / " + total);
            }
            else {
                progress.html("地址定位作業完成, 共 " + total + " 筆 (成功 " + success + " 筆, 失敗 " + failed + " 筆)");
                locateOnGoing = false;
            }
        }

        function geoLocator() {

            var qryULs = $("ul[id^=lv_classQueryResult]");
            total = qryULs.length;
            success = 0;
            failed = 0;
            i = 0;
            locateOnGoing = true;
            addrList = [];

            qryULs.each(function () {

                refreshProgress();

                if ($(this).data("checked") != 'Y') { return; }

                var addr = {};
                addr.ocid = $(this).data("id");
                addr.spAddr = $(this).find("span[id=spTAddress]").text();
                addr.wgs84_y = $(this).find("span[id=WGS84]").find("[id=Y]");
                addr.wgs84_x = $(this).find("span[id=WGS84]").find("[id=X]");

                if (addr.wgs84_y.val() != "" && addr.wgs84_y.val() != "0") {
                    console.log(addr.ocid + " has WGS84 already, skip.");
                    return;
                }

                addrList.push(addr);
            });

            total = addrList.length;

            checkRunLocate();
        }

        function checkRunLocate() {
            while (locateProcess < 5 && i < addrList.length) {
                locateProcess++;
                doLocate(addrList[i]);
                i++;
            };

            if (i < addrList.length) {
                setTimeout(function () {
                    checkRunLocate();
                }, 100);
            }
        }

        function doLocate(addr) {
            addr.wgs84_y.addClass("red").val("0");
            addr.wgs84_x.addClass("red").val("0");

            locateSearch('', '', addr.spAddr, undefined,
                function (reqStr, env, loc) {
                    // success, loc: [Y, X]
                    addr.wgs84_y.val(loc[0]).removeClass("red").addClass("green");
                    addr.wgs84_x.val(loc[1]).removeClass("red").addClass("green");
                    success++;
                    locateProcess--;
                    refreshProgress();
                },
                function () {
                    // failed
                    console.log(addr.spAddr + " .... geoLocator failed");
                    failed++;
                    locateProcess--;
                    refreshProgress();
                },
            true, true);
        }

        function SubOCIDAddress(ocid1, address1) {
            $('#txtocid2').val(ocid1);
            $('#txtaddress2').val(address1);
        }
    </script>

    <style type="text/css">
        input, button { padding: 5px; border-radius: 2px; border: 1px solid #AAA; }
        input { text-align: right; }
        .red { color: red; }
        .green { color: red; }
    </style>

</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;TGOS地址定位班級</asp:Label>
                </td>
            </tr>
        </table>
        <div id="schclass1">
            <dl>
                <dt>查詢筆數</dt>
                <dd>
                    <asp:TextBox ID="txtROWNUM" runat="server" MaxLength="5" Columns="6">6</asp:TextBox></dd>
                <dt>班級流水號OCID</dt>
                <dd>
                    <asp:TextBox ID="txtocid2" runat="server" MaxLength="10" Columns="6"></asp:TextBox></dd>
                <dt>調整地址</dt>
                <dd>
                    <asp:TextBox ID="txtaddress2" runat="server" MaxLength="300" Columns="50"></asp:TextBox></dd>
            </dl>
            <asp:Button ID="btnSch2" runat="server" Text="重設地址" />
            <asp:Button ID="btnSch3" runat="server" Text="重新查詢" />
            <input id="txt_localareanet" runat="server" />
            <p></p>
        </div>
        <div>
            <div id="classQueryDiv1" runat="server">
                <asp:Literal ID="litMsg" runat="server"></asp:Literal>

                <button type="button" id="btnCheckAll">全選</button>
                &nbsp;&nbsp;
                <button type="button" id="btnUncheckAll">全不選</button>
                &nbsp;&nbsp;
                <button type="button" id="btnLocate">TGos地址定位</button>
                &nbsp;&nbsp;
                <span id="litProgress" class="red"></span>
                <br />
                <button type="button" id="btnSave">儲存定位座標</button>
                <span id="litSaveProgress" class="red"></span>
                <br />

                <asp:ListView ID="lv_classQueryResult" runat="server">
                    <ItemTemplate>
                        <ul runat="server" id="classItemUL" data-checked="Y">
                            <li>
                                <span id="spOCID"><asp:Literal ID="litOCID" runat="server"></asp:Literal></span>
                                <input type="checkbox" checked="checked" />
                            </li>
                            <li>
                                <span id="spClsName"><asp:Literal ID="litClassCName" runat="server"></asp:Literal></span>
                            </li>
                            <li>
                                <span id="spTAddress"><asp:Literal ID="litTAddress" runat="server"></asp:Literal></span>
                                &nbsp;&nbsp;&nbsp;&nbsp;【 <span id="WGS84">WGS84座標：<input id="Y" placeholder="Y" />,<input id="X" placeholder="X" />】</span>
                                <button type="button" id="btnselect1" runat="server">選</button>
                            </li>
                        </ul>
                    </ItemTemplate>
                    <EmptyDataTemplate>
                        <div style="text-align: center; margin-top: 15px;">
                            <span style="color: red">無待地址定位！</span>
                        </div>
                    </EmptyDataTemplate>
                </asp:ListView>

            </div>
        </div>
    </form>
</body>
</html>
