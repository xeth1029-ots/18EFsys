<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_002_R.aspx.vb" Inherits="WDAIIP.TC_01_002_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練機構明細表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" src="../../js/common.js"></script>
    <script language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript">
        function ReportPrint() {
            if (document.form1.yearlist.value == '' || document.form1.planlist.value == '') {
                alert('請選擇年度及訓練計畫');
                return false;
            }
            return true;
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="0" cellpadding="0" width="100%" border="0">
            <%--<tr>
			<td>
				<table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
                            首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000">列印訓練機構明細</font>
							</asp:Label>
							<font color="#990000">&nbsp;</font> </td>
					</tr>
				</table>
			</td>
		</tr>--%>
            <tr>
                <td>
                    <table class="table_sch" width="100%" cellpadding="1" cellspacing="1">
                        <tr>
                            <td class="bluecol_need" width="20%">年度 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="yearlist" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">訓練計畫 </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="planlist" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="DistrictList" runat="server">
                                </asp:DropDownList>
                            </td>
                            <td class="bluecol">縣市 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TBCity" runat="server" onfocus="this.blur()" Width="50%"></asp:TextBox>
                                <input id="city_zip" onclick="getZip('../../js/Openwin/zipcode_search.aspx', 'TBCity', 'city_code', 'city_code')" type="button" value="..." name="city_zip" runat="server">
                                <input id="city_code" type="hidden" name="city_code" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" width="20%">機構名稱 </td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="OrgName" runat="server" Columns="40" Width="70%"></asp:TextBox>
                            </td>
                            <td class="bluecol" width="20%">統編 </td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="ComIDNO" runat="server" Width="40%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">機構別 </td>
                            <td colspan="3" class="whitecol">
                                <asp:DropDownList ID="OrgTypeList" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Button ID="btnPrint" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
