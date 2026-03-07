<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_011_R.aspx.vb" Inherits="WDAIIP.TC_01_011_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>列印訓練機構郵遞標籤</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function ReportPrint() {
            if (document.form1.yearlist.value == '' || document.form1.planlist.value == '') {
                alert('請選擇年度及訓練計畫');
                return false;
            }
            return true;
        }
    </script>
    <%--<script type="text/javascript" src="../../js/jquery-1.6.2.js"></script>
    <script type="text/javascript" src="../../js/selectControl.js.aspx" charset="UTF-8"></script>--%>
    <%--<script language="javascript">
            function yearPlan(selectedPlanID) {
                var year = document.getElementById('yearlist');
                var parms = "[['year','" + year.value + "']]";      // 透過 selectControl 傳遞給 SQLMap 的年度查詢條件, 格式請參考 selectControl 定義說明
                selectControl('ajaxTPlanList', 'planlist', 'PlanName', 'TPlanID', '請選擇', selectedPlanID, parms);
            }
        </script>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
                <td class="font">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000">列印訓練機構郵遞標籤</font>
                </td>
            </tr>
            <tr>
                <td>
                    <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                        <tr>
                            <td width="100" class="bluecol_need">年度
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="yearlist" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訓練計畫
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="planlist" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="100">轄區
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DistrictList" runat="server">
                                </asp:DropDownList>
                            </td>
                            <td class="bluecol" width="100">縣市
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBCity" runat="server" onfocus="this.blur()" Columns="26"></asp:TextBox>
                                <input id="city_zip" onclick="getZip('../../js/Openwin/zipcode_search.aspx', 'TBCity', 'city_code', 'city_code')" type="button" value="..." name="city_zip" runat="server" class="button_b_Mini">
                                <input id="city_code" style="width: 26px; height: 22px" type="hidden" name="city_code" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">機構名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="OrgName" runat="server" Columns="30" MaxLength="30"></asp:TextBox>
                            </td>
                            <td class="bluecol">統編
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="ComIDNO" runat="server" Columns="15" MaxLength="10"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">機構別
                            </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="OrgTypeList" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%" border="0">
                        <tr>
                            <td class="whitecol">
                                <div align="center">
                                    <asp:Button ID="btnPrint" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>&nbsp;
                                </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
