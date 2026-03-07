<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_002.aspx.vb" Inherits="WDAIIP.TC_01_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練機構設定</title>
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
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
        function but_edit(orgid, planid, rid, distid, id) {
            location.href = 'TC_01_002_add.aspx?orgid=' + orgid + '&planid=' + planid + '&rid=' + rid + '&ProcessType=Update&distid=' + distid + '&ID=' + id;
        }

        function but_del(orgid, account, classid, rid, planid, is_parent, id) {
            if (is_parent) {
                alert("此機構尚有下層單位,不可刪除!!");
                return;
            }
            if (classid == "" && account == "") {
                if (window.confirm("此動作會刪除機構資料，是否確定刪除?")) {
                    location.href = 'TC_01_002_del.aspx?orgid=' + orgid + '&rid=' + rid + '&planid=' + planid + '&ID=' + id;
                }
            } else if (classid != "") {
                alert('此機構已有開班資料，不可以刪除!!');
            } else {
                alert('此機構已有帳號資料，不可以刪除!!');
            }
        }

        function Search() {
            //debugger;
            var IsApply = document.getElementById('IsApply');
            var rblOrgLevel = document.getElementById('rblOrgLevel');
            var city_code = document.getElementById('city_code');
            var TB_ComIDNO = document.getElementById('TB_ComIDNO');
            var TB_OrgName = document.getElementById('TB_OrgName');
            var hidLID = document.getElementById("hidLID");
            //if (hidLID.value == "0" || hidLID.value == "1") { //2018 改版修正(配合 mark 資料狀態查詢條件的作法，改判斷登入者層級)
            //if (IsApply.children(0).checked == true) {
            if (!IsApply) { return; }
            if (IsApply.children[0].checked == true) {
                if (!rblOrgLevel) { return; }
                var rblOrgLevelval = getRadioButtonListValue(rblOrgLevel);
                //alert(rblOrgLevelval);
                if (rblOrgLevelval == 'N' && city_code.value == '' && TB_ComIDNO.value == '' && TB_OrgName.value == '') {
                    //alert('請輸入機構名稱、統一編號或者是縣市代碼');
                    blockAlert("請輸入機構名稱、統一編號或者是縣市代碼", "提示訊息");
                    return false;
                }
            }
        }

        function but_share(orgid, planid, rid, distid, id) {
            location.href = 'TC_01_002_add.aspx?orgid=' + orgid + '&ProcessType=Share&distid=' + distid + '&planid=' + planid + '&rid=' + rid + '&ID=' + id;
        }

        // 計畫別下拉選項若選產投 or 充飛計畫，才能顯示(機構屬性)機構別的查詢條件
        function chgPlan() {
            var obj = document.getElementById("drpPlan");
            var planid = obj.options[obj.selectedIndex].value;
            var trPlanPoint = document.getElementById("trPlanPoint");
            var rblPlanPoint = document.getElementById("rblPlanPoint");
            var dl_typeid2 = document.getElementById("dl_typeid2");
            if (planid == "28" || planid == "54") {
                trPlanPoint.style.display = "";
            }
            else {
                trPlanPoint.style.display = "none";
                //重設 機構屬性選項值
                if (!rblPlanPoint) { return; }
                rblPlanPoint.children[0].checked = true;
                dl_typeid2.options.length = 0;
            }
        }

        function IsApply_display(obj) {
            //debugger;
            var DISTTR = document.getElementById('DISTTR');
            var orglevelTR = document.getElementById('orglevelTR');
            if (!DISTTR) { return; }
            if (!orglevelTR) { return; }
            DISTTR.style.display = 'none';
            orglevelTR.style.display = 'none';
            if (getRadioButtonListValue(obj) == 'Y') {
                //DISTTR.style.display = 'inline';
                //orglevelTR.style.display = 'inline';
                DISTTR.style.display = '';
                orglevelTR.style.display = '';
            }
        }

        function getRadioButtonListValue(obj) {
            var result = "";
            for (var i = 0; i < obj.childNodes.length; i++) {
                if (obj.childNodes[i].checked) {
                    result = obj.childNodes[i].value;
                    break;
                }
            }
            return result;
        }
    </script>
    <%--
	<script type="text/javascript" src="../../js/selectControl.js.aspx" charset="UTF-8"></script>
    <script language="javascript">
    function yearPlan(selectedPlanID) {
        var year = document.getElementById('Years');
        var parms = "[['year','" + year.value + "']]";      // 透過 selectControl 傳遞給 SQLMap 的年度查詢條件, 格式請參考 selectControl 定義說明
        selectControl('ajaxTPlanList', 'drpPlan', 'PlanName', 'TPlanID', '請選擇', selectedPlanID, parms);
    }
    </script>
    --%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;訓練機構設定</asp:Label>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="hidLID" runat="server" />
        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td id="td6" runat="server" class="bluecol" width="20%">機構名稱 </td>
                <td class="whitecol" width="30%">
                    <asp:TextBox ID="TB_OrgName" runat="server" Columns="40" Width="80%"></asp:TextBox></td>
                <td id="td7" runat="server" class="bluecol" width="20%">統編 </td>
                <td class="whitecol" width="30%">
                    <asp:TextBox ID="TB_ComIDNO" runat="server" Columns="10" MaxLength="10" Width="60%"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">縣市 </td>
                <td colspan="3" class="whitecol">
                    <asp:TextBox ID="TBCity" runat="server" Columns="30" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <input id="city_zip" onclick="getZip('../../js/Openwin/zipcode_search.aspx', 'TBCity', 'zip_code', 'city_code')" type="button" value="..." name="city_zip" runat="server" class="button_b_Mini">
                </td>
            </tr>
            <tr id="DISTTR">
                <td class="bluecol">轄區分署 </td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="DistID" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol">機構別 </td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="OrgKindList" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol">年度 </td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="Yearlist" runat="server"></asp:DropDownList></td>
            </tr>
            <tr>
                <td class="bluecol">計畫別 </td>
                <td colspan="3" class="whitecol">
                    <asp:DropDownList ID="drpPlan" runat="server"></asp:DropDownList></td>
            </tr>
            <%-- 2018 新增：配合「訓練機構屬性設定」功能整併到「訓練機構設定」功能，加入「(機構屬性)機構別」查詢條（計畫別選產投時此查詢條件欄位才顯示）--%>
            <tr id="trPlanPoint" runat="server" style="display: none">
                <td class="bluecol">(機構屬性)機構別 </td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="rblPlanPoint" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" AutoPostBack="True">
                        <asp:ListItem Value="0" Selected="True">不區分</asp:ListItem>
                        <asp:ListItem Value="1">產業人才投資計畫</asp:ListItem>
                        <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                    </asp:RadioButtonList><br />
                    <asp:DropDownList ID="dl_typeid2" runat="server"></asp:DropDownList>
                </td>
            </tr>
            <%-- 2018 先 mark(待確認資料邏輯)： 審核中的資料 org_apply 沒有欄位記錄 (機構屬性)機構別 資料，--%>
            <tr id="tr_IsApply" runat="server" style="display: none">
                <td class="bluecol">資料狀態 </td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="IsApply" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="Y" Selected="True">正式</asp:ListItem>
                        <asp:ListItem Value="N">審核中</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr id="orglevelTR" runat="server">
                <td class="bluecol">機構層級 </td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="rblOrgLevel" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="N" Selected="True">不區分</asp:ListItem>
                        <asp:ListItem Value="2">2</asp:ListItem>
                        <asp:ListItem Value="3">3</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <%--style="display: none"--%>
            <tr style="display: none">
                <td class="bluecol">匯入機構評鑑 </td>
                <td colspan="3" class="whitecol">
                    <div align="left">
                        <input id="File1" type="file" size="40" name="File1" runat="server" accept=".xls,.ods" />
                        <asp:Button ID="Btn_XlsImport" runat="server" Text="匯入評鑑" CssClass="asp_button_M"></asp:Button>
                        <%--<asp:button id="Btn_Import" runat="server" Text="匯入名冊"></asp:button>,(必須為csv格式),<asp:hyperlink id="Hyperlink2" runat="server" ForeColor="#8080FF" CssClass="font" NavigateUrl="../../Doc/Org_Import.zip">下載整批上載格式檔</asp:hyperlink>,--%>
					(必須為ods或xls格式)<asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" NavigateUrl="../../Doc/Org_Comments.zip" ForeColor="#8080FF">下載整批上載格式檔</asp:HyperLink>&nbsp;&nbsp;
                    </div>
                </td>
            </tr>
            <tr>
                <td class="bluecol">匯出檔案格式</td>
                <td colspan="3" class="whitecol">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td class="whitecol">
                    <asp:DataGrid ID="dtgAddresses1" GridLines="Both" CellPadding="8" HeaderStyle-CssClass="head_navy" ItemStyle-CssClass="font" runat="server" Width="100%">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td class="whitecol" align="center">
                    <input id="city_code" style="width: 26px; height: 22px" type="hidden" name="city_code" runat="server">
                    <asp:Button ID="bt_search" Text="查詢" runat="server" CssClass="asp_button_S"></asp:Button>
                    <asp:Button ID="bt_add" Text="新增" runat="server" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="bt_EXPORT" runat="server" Text="匯出" Visible="False" CssClass="asp_Export_M"></asp:Button>
                    <%--
                    <input id="check_add" style="width: 40px; height: 22px" type="hidden" name="check_add" runat="server">
				    <input id="check_del" style="width: 48px; height: 22px" type="hidden" size="2" name="check_del" runat="server">
				    <input id="check_mod" style="width: 45px; height: 22px" type="hidden" size="2" name="check_mod" runat="server">
                    --%>
                    <input id="zip_code" type="hidden" size="2" name="zip_code" runat="server">
                    <div align="center"></div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
        </table>
        <br>
        <asp:Panel ID="Panel" runat="server" Width="100%" Visible="False">
            <asp:DataGrid ID="DG_Org" runat="server" CssClass="font" Width="100%" Visible="False" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                <Columns>
                    <asp:BoundColumn HeaderText="編號">
                        <HeaderStyle HorizontalAlign="Center" Width="2%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <%--<asp:BoundColumn DataField="name" HeaderText="轄區中心">--%>
                    <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區分署">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="PlanName" HeaderText="計畫名稱">
                        <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ComIDNO" HeaderText="統編">
                        <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="Address" HeaderText="地址">
                        <HeaderStyle HorizontalAlign="Center" Width="12%"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ActNo" HeaderText="保險證號">
                        <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ContactName" HeaderText="聯絡人姓名">
                        <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ContactEmail" HeaderText="聯絡人E-Mail">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                    </asp:BoundColumn>
                    <asp:TemplateColumn HeaderText="功能">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                        <ItemTemplate>
                            <asp:LinkButton ID="lbtShare" runat="server" Text="共用" CommandName="share" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="lbtEdit" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="lbtDel" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="lbtChk" runat="server" Text="審核" CommandName="" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="lbtYear" runat="server" Text="年度對應功能" CommandName="year" CssClass="linkbutton" Visible="false"></asp:LinkButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <PagerStyle Visible="False"></PagerStyle>
            </asp:DataGrid>
            <%--<font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>--%>
            <div align="center">
                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
            </div>
        </asp:Panel>
    </form>
</body>
</html>
