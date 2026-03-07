<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_008.aspx.vb" Inherits="WDAIIP.SD_03_008" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員電子郵件匯出</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button4').click();
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value + '&amp;PlanID=' + document.form1.PlanID.value);
        }

        //function choose_other() {
        //    var OCID = document.form1.OCIDValue1.value
        //    window.open('SD_02_003_other.aspx?OCID=' + OCID, '', 'width=550,height=250,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
        //}

        //function search(){
        //	if(document.form1.OCIDValue1.value==''){
        //		alert('請選擇職類班別!')
        //		return false;
        //	}
        //}
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;學員電子郵件匯出</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>

                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol_need" width="100">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="Syear" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="100">計畫名稱
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="PlanID" runat="server" Width="420px">
                                </asp:DropDownList>
                                <input id="DistID" type="hidden" name="DistID" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol_need">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="Button8" onclick="javascript: wopen('../../Common/LevOrg2.aspx?DistID=' + document.form1.DistID.value + '&amp;PlanID=' + document.form1.PlanID.value, '訓練機構', 400, 400, 1)" type="button" value="..." name="Button8" runat="server" class="asp_button_Mini" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button4"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="312px">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">
                                <asp:Label ID="LabCJOB_UNKEY" runat="server">
											通俗職類</asp:Label>
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="50%"></asp:TextBox>
                                <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                                <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol">匯出對象
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="TYPE1" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="1" Selected="True">不區分</asp:ListItem>
                                    <asp:ListItem Value="2">在訓</asp:ListItem>
                                    <asp:ListItem Value="3">結訓</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol">匯出格式
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="TYPE2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="1" Selected="True">Outlook格式</asp:ListItem>
                                    <asp:ListItem Value="2">Outlook Express格式</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Button ID="Button1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
