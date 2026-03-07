<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_034.aspx.vb" Inherits="WDAIIP.SD_15_034" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>重大災害數據統計</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript"></script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;重大災害數據統計</asp:Label>
                </td>
            </tr>
        </table>
        <div id="div_search1" runat="server">
            <table class="table_sch" width="100%" border="0">
                <tr>
                    <td class="bluecol_need" width="20%">重大災害名稱</td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="DDL_DISASTER_S1" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">轄區分署</td>
                    <td class="whitecol" colspan="3">
                        <asp:DropDownList ID="DDL_DISTID_S1" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">訓練計畫</td>
                    <td class="whitecol" colspan="3">
                        <asp:CheckBoxList ID="CBL_TPLANID_S1" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                        </asp:CheckBoxList>
                        <asp:Label ID="LabTPLANNAME" runat="server" Text=""></asp:Label>
                        <input id="CBL_TPLANID_S1Hidden" type="hidden" value="0" name="CBL_TPLANID_S1Hidden" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">資料篩選方式</td>
                    <td class="whitecol" colspan="3">
                        <span>使用
                            <asp:RadioButtonList ID="RBL_ETENTER_USE" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                <asp:ListItem Value="1" Selected="True">系統預設區間</asp:ListItem>
                                <asp:ListItem Value="2">自選日期區間</asp:ListItem>
                            </asp:RadioButtonList></span><br />
                        <span style="color: red">1.系統預設區間：<br />
                            報名人數：【開訓日】：(災害起始日＋1個月)~(今日＋30天)，【報名日】&gt;=災害起始日<br />
                            參訓人數：【開訓日】： 災害起始日~今日<br />
                            <br />
                            2.自選日期區間：使用者自行輸入日期
                        </span>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">報名日期區間 </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="ETENTERDATE1" runat="server" Columns="14" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ETENTERDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
					<asp:TextBox ID="ETENTERDATE2" runat="server" Columns="14" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('ETENTERDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">開訓日期期間 </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="STDATE1" runat="server" Columns="14" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDATE1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
					<asp:TextBox ID="STDATE2" runat="server" Columns="14" MaxLength="10"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDATE2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">匯出明細範圍 </td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="RBL_SCOPE1" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="01">報名人數</asp:ListItem>
                            <asp:ListItem Value="02" Selected="True">參訓人數</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>

                <tr>
                    <td class="whitecol" align="center" colspan="4">
                        <asp:Button ID="BTN_SEARCH1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        &nbsp;<asp:Button ID="BTN_EXPORT1" runat="server" Text="匯出明細" CssClass="asp_Export_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </div>
        <div id="div_detail1" runat="server">
            <table id="Table1_view" runat="server" class="table_sch" cellspacing="0" cellpadding="6" border="1" style="width: 100%; border-collapse: collapse;">
                <tr>
                    <td class="table_title_left" colspan="4">
                        <asp:Label ID="LAB_TITLE1" runat="server" Text="重大災害名稱：(查無資料)"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center" style="width: 20%;">訓練計畫</td>
                    <td align="left" colspan="3">
                        <asp:Label ID="LAB_TPLANNAME" runat="server" Text="(查無資料)"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center" style="background-color: WhiteSmoke;">轄區分署</td>
                    <td align="left" style="background-color: WhiteSmoke;" colspan="3">
                        <asp:Label ID="LAB_DISTNAME" runat="server" Text="(查無資料)"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center">受災地區</td>
                    <td align="left" colspan="3">
                        <asp:Label ID="LAB_AREAS" runat="server" Text="(查無資料)"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center" style="background-color: WhiteSmoke;">回報日期</td>
                    <td align="left" style="background-color: WhiteSmoke;" colspan="3">
                        <asp:Label ID="Lab_ReturnDate" runat="server" Text="(查無資料)"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center">資料篩選方式</td>
                    <td align="left" colspan="3">
                        <asp:Label ID="LabETENTER" runat="server" Text="(自選日期區間)"></asp:Label></td>
                </tr>
                <tr id="tr_SHOWMSG_1a" runat="server">
                    <td align="center">報名日期區間</td>
                    <td align="left" colspan="3">
                        <asp:Label ID="Lab_RegDateRange" runat="server" Text="(查無資料)"></asp:Label></td>
                </tr>
                <tr id="tr_SHOWMSG_1b" runat="server">
                    <td align="center" style="background-color: WhiteSmoke;">開訓日期區間</td>
                    <td align="left" style="background-color: WhiteSmoke;" colspan="3">
                        <asp:Label ID="Lab_TrainDateRange" runat="server" Text="(查無資料)"></asp:Label></td>
                </tr>
                <tr>
                    <td align="center">報名人數</td>
                    <td align="left"><font color="red" size="4">
                        <asp:Label ID="Lab_NumberAPP" runat="server" Text="(查無資料)"></asp:Label></font>
                    </td>
                    <td colspan="2">
                        <asp:Label ID="Lab_APPMSG" runat="server" Text=""></asp:Label></td>
                </tr>
                <tr id="tr_SHOWMSG_1d" runat="server">
                    <td align="center" style="background-color: WhiteSmoke;">參訓人數</td>
                    <td align="left" style="background-color: WhiteSmoke;"><font color="red" size="4">
                        <asp:Label ID="Lab_NumberPart" runat="server" Text="(查無資料)"></asp:Label></font></td>
                    <td colspan="2" style="background-color: WhiteSmoke;">
                        <asp:Label ID="Lab_PartMSG" runat="server" Text=""></asp:Label></td>
                </tr>
                <tr>
                    <td align="center" colspan="4">
                        <asp:Button ID="BtnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </div>
        <asp:HiddenField ID="hid_ADID" runat="server" />
        <asp:HiddenField ID="hid_ADID_ZIPCODES" runat="server" />
    </form>
</body>
</html>
