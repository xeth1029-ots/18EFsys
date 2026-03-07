<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_008.aspx.vb" Inherits="WDAIIP.SD_15_008" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓後動態調查統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button5').click();
        }
        function print() {
            var msg = '';
            //Button1 //if(document.getElementById('OCIDValue1').value=='') msg+='請選擇班級\n';
            if (document.getElementById('yearlist').selectedIndex == 0) msg += '請選擇年度\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function choose_class() {
            var Hid_LID = document.getElementById('Hid_LID');
            var Hid_YEARS = document.getElementById('Hid_YEARS');
            var RIDValue = document.getElementById('RIDValue');

            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';

            var s_oclass = '../02/SD_02_ch.aspx?&RID=' + RIDValue.value;
            if ((Hid_YEARS.value != "") && (Hid_LID.value == "0" || Hid_LID.value == "1")) {
                s_oclass = '../02/SD_02_ch.aspx?selected_year=' + Hid_YEARS.value + '&RID=' + RIDValue.value;
            }
            openClass(s_oclass);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;訓後動態調查統計表</asp:Label>
                </td>
            </tr>
        </table>
        <%--<input id="Years" type="hidden" name="Years" runat="server">--%>
        <input id="SOCIDValue" type="hidden" name="SOCIDValue" runat="server">
        <%--880FE13F7EEC3837AEF83751BB45F885B99875DC01F97A4B619CB2B03D92B78A20AE1C162D9A43DBE0E21E22D859F42DCA92A74C7B959CE8EC9D58034AF7742
            FD51FB3367753C6E4F9AF0D2E2F7B78D67AA59B030710B1C1A91E9A040F4BC45F88DF0F720C0B89A6033594E702ED8343B8873C0065C02A2461CFAF058E0BAC
            17D5E9F72FA8F0D2AF1DA5A851AFA62C8B5900D973E4648D298D348377FC684A9990423018CB99DA16CE3CD0583774E54B2B10375C70C5152FFA11F2B7CEAF7
            ABC543E91060A406CDD30F8BB9DA1346A1601DD93EF26942DB65A636994084D9A6CB39ADB58496465179FF569111CCF0ECB4C7B193AE0790E8EC71B7C31FFDF
            21DB4989A5848482FE42C2205CE4F0CBBBB6AE4038C4E50F5291698F1CBD7041274AB0C07E26F9CA5D106452A991976729CA6453BFB970E836B4144FE0D2EBC
            7FF96FAA89838F5E16E7495270D6C4A68F5A7CCD62534B86F312025BB8165331B01E6--%>
        <table class="font" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td>
                    <table class="table_sch" width="100%">
                        <tr id="tr_yearlist" runat="server">
                            <td class="bluecol_need" style="width: 20%">年度 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                                <asp:Label ID="Label1" runat="server" Font-Bold="True" ForeColor="Red" Text=" (匯出助益率分析表)年度不選時，開訓期間為必填!"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="55%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button5" Style="display: none" runat="server" Text="Button5"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server"  Width="25%"></asp:TextBox><%--onfocus="this.blur()"--%>
                                <asp:TextBox ID="OCID1" runat="server"  Width="30%"></asp:TextBox><%--onfocus="this.blur()"--%>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span></td>
                        </tr>
                         <tr id="mode1_1">
                            <td class="bluecol">開訓期間
                            </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">～
								<asp:TextBox ID="STDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30">
                            </td>
                        </tr>
                        <tr id="trPlanKind" runat="server">
                            <td class="bluecol">計畫範圍 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="SearchPlan" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trPackageType" runat="server">
                            <td class="bluecol">包班種類 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PackageType" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="A" Selected="True">全部</asp:ListItem>
                                    <%--<asp:ListItem Value="1">非包班</asp:ListItem>--%>
                                    <asp:ListItem Value="2">企業包班</asp:ListItem>
                                    <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">交叉查詢選項 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="FuncID" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="BtnPrint3" runat="server" Text="列印明細" Visible="False" CssClass="asp_Export_M"></asp:Button>
                                <asp:Button ID="BtnExport1" runat="server" Text="匯出助益率分析表"   CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <%--2EAC7325FB5EA1DF0F0EAF19E68189E73317D048D7332CC0E58C90C9BA8A40EBC22645E24A771936C5CC6FDA1958FAB3A5121B4DFCD487C05C639D08B2B598CE56D71
            DD10A3190D64D60CFD3EC6E33D8BA1A667BA82BCE81AA8F6810EB284B8F3E77D53EA73222E48E09A1C2333C04A3D0E09611357082F0B1DD41D7D0943091494B8AFF8C
            5B6704807B013A145F1D1C95DD4F145C97D1F9ABF7EBD7C703D984DC3D397826EE5EBE14F05F5157B7C6C06895E52688483D2028193C0412305E41FBA19C0E172FB80
            DC429379C285056CAFF55FF620AEABED7FCF4AD759DE67D2CB220B5C3C6C7450B87FB57EFE3B1DCE06EFE68825A4B90DBB68586DC15D594D8799B09BB8E65C82F3F01
            5D59A4FDEE26AF14DED47C0F80704BFCD900884A640497DC108E8FF4D32281266688C482505F28E6EC8F1DB11F7D6F8E8BE830486ADBBFEF8566EF080970013CAF916
            CB724CFDE76FF83FE2377B918B1462C02E44322786CF28E1676916EDFE26E5A8B062923413C84A7C0A6A75BCA9E8F6D79EEA7F0D48D12028A187B72FAD6DB72A79712
            6F8611F2A2978507A9DDFC1AF726D3E4B2B80429410C5166FAD5409D6BA14E1D48BC1E1307C98578512B7C20C6A22317C5AEC1238E16D8F0E7AFD524EB7E99AF775A0
            007BEC00534C1BEB6E91EBE49246E--%>
        <asp:HiddenField ID="Hid_LID" runat="server" />
        <asp:HiddenField ID="Hid_YEARS" runat="server" />
    </form>
</body>
</html>
