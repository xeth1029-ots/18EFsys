<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_017_add.aspx.vb" Inherits="WDAIIP.TC_01_017_add" %>

<html>
<head>
    <title>訓練機構屬性設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //檢查zipcode(City欄位名,Zip欄位名,Zip輸入內容)
        function getZipName(CityID, ZipID, ZipValue) {
            if (!isBlank(ZipID)) {
                if (isUnsignedInt(ZipValue) && ZipValue.length == 3) {
                    ifmCheckZip.document.form1.hidCityID.value = CityID;
                    ifmCheckZip.document.form1.hidZipID.value = ZipID.id;
                    ifmCheckZip.document.form1.hidValue.value = ZipValue;
                    ifmCheckZip.document.form1.submit();
                } else {
                    ZipID.value = '';
                    document.getElementById(CityID).value = '';
                    ZipID.focus();
                    alert('查無' + ZipValue + '郵遞區號!');
                }
            } else {
                document.getElementById(CityID).value = '';
            }
        }

        //判斷儲存
        function chkSave() {
            var msg = '';
            var dl_typeid1 = document.getElementById('dl_typeid1');
            var dl_typeid2 = document.getElementById('dl_typeid2');
            var city_code = document.getElementById('city_code');
            var ZipCODEB3 = document.getElementById('ZipCODEB3');
            var TBaddress = document.getElementById('TBaddress');

            if (dl_typeid1.value == '') msg += '請選擇計畫別!\n';
            if (dl_typeid2.value == '') msg += '請選擇機構別!\n';
            if (isBlank(city_code)) msg += '請輸入郵遞區號前3碼!\n';
            if (isBlank(ZipCODEB3)) msg += '請輸入郵遞區號後2碼!\n';
            if (isBlank(TBaddress)) msg += '請輸入地址!\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
    <%--<script type="text/javascript" src="../../js/jquery-1.6.2.js"></script>
	<script type="text/javascript" src="../../js/selectControl.js.aspx" charset="UTF-8"></script>--%>
    <%--<script language="javascript">
		    function SelPlan(selectedID) {
		        var selValue = '0';
		        var dl_typeid1 = document.getElementById('dl_typeid1');
		        selValue = dl_typeid1.value;
		        if (selValue != '') {
		            var parms = "[['TypeID1','" + selValue + "']]";      // 透過 selectControl 傳遞給 SQLMap 的年度查詢條件, 格式請參考 selectControl 定義說明
		            selectControl('ajaxQueryKeyOrgTypeS2', 'dl_typeid2', 'TypeID2Name', 'TypeID2', '請選擇', selectedID, parms);
		        }
		        else {
		            var obj = document.getElementById('dl_typeid2');
		            obj.length = 1;
		        }
		    }
		</script>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table2" width="100%">
            <tr>
                <td class="font">
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server"> 首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000">訓練機構屬性設定</font> </asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table class="table_nw" width="100%">
            <tr>
                <td width="15%" class="bluecol">機構名稱
                </td>
                <td width="85%" class="whitecol" colspan="3">
                    <asp:TextBox ID="tb_orgname" runat="server" Width="410px" onfocus="this.blur()"></asp:TextBox>
                    <input id="hid_orgid" type="hidden" name="hid_orgid" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="bluecol">統一編號
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="tb_comidno" runat="server" Width="88px" onfocus="this.blur()" MaxLength="10"></asp:TextBox>
                </td>
            </tr>
            <tr>
                <td class="bluecol">計畫別
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="dl_typeid1" runat="server" Width="200px" AutoPostBack="True">
                        <asp:ListItem Value="">==請選擇==</asp:ListItem>
                        <asp:ListItem Value="1">產業人才投資計畫</asp:ListItem>
                        <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">機構別
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="dl_typeid2" runat="server" Width="200px">
                        <asp:ListItem Value="">==請選擇==</asp:ListItem>
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">立案地址/會址
                </td>
                <td class="whitecol">
                    <input id="city_code" maxlength="3" name="city_code" runat="server" />－
                    <input id="ZipCODEB3" maxlength="3" runat="server" />
                    <input id="hidZipCODE6W" type="hidden" runat="server" />
                    <asp:Literal ID="Litcity_code" runat="server"></asp:Literal><%--郵遞區號查詢--%>
                    <br />
                    <asp:TextBox ID="TBCity" runat="server" Width="130px" onfocus="this.blur()"></asp:TextBox>
                    <input id="city_zip" type="button" value="..." name="city_zip" runat="server" class="button_b_Mini" />
                    <asp:TextBox ID="TBaddress" runat="server" Width="60%"></asp:TextBox>
                    <asp:RequiredFieldValidator ID="rfvcity" runat="server" ErrorMessage="請選擇縣市" Display="None" ControlToValidate="TBCity"></asp:RequiredFieldValidator>
                    <asp:RequiredFieldValidator ID="rfvaddress" runat="server" ErrorMessage="請輸入地址" Display="None" ControlToValidate="TBaddress"></asp:RequiredFieldValidator>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" class="whitecol">
                    <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                    <asp:Button ID="bt_back" runat="server" Text="回上一頁" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                    <br>
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <iframe id="ifmCheckZip" name="ifmCheckZip" src="../../common/CheckZip.aspx" width="0%" height="0%"></iframe>
    </form>
</body>
</html>
