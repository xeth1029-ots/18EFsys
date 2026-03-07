<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_007.aspx.vb" Inherits="WDAIIP.SD_15_007" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓後意見調查統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
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
            //if(document.getElementById('OCIDValue1').value=='') msg+='請選擇班級\n';
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;訓後意見調查統計表</asp:Label>
                </td>
            </tr>
        </table>
        <input id="Years" type="hidden" name="Years" runat="server">
        <input id="SOCIDValue" type="hidden" name="SOCIDValue" runat="server">
        <%--<table class="font" id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0"><tr><td>
    <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
    首頁&gt;&gt;學員動態管理&gt;&gt;統計表(產學訓表單列印)&gt;&gt;<FONT color="#990000">訓後意見調查統計表</FONT></asp:Label></td></tr></table><br />--%>
        <table class="font" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td>
                    <table class="table_sch" width="100%">
                        <tr id="tr_yearlist" runat="server">
                            <td class="bluecol_need" style="width: 20%">年度 </td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="True"></asp:DropDownList>
                                <asp:Label ID="lab_yearlist" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="55%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" value="..." type="button" name="Button2" runat="server" class="button_b_Mini">
                                <asp:Button Style="display: none" ID="Button5" runat="server" Text="Button5"></asp:Button>
                                <span style="position: absolute; display: none" id="HistoryList2" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" value="..." type="button" class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span style="position: absolute; display: none; left: 270px" id="HistoryList">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="trPlanKind" runat="server">
                            <td class="bluecol">計畫範圍 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="SearchPlan" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trSTDate12" runat="server">
                            <td class="bluecol" width="20%">開訓期間 </td>
                            <td class="whitecol" colspan="3">
                                <span id="span01" runat="server">
                                    <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
							    <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                    (配合列印其他意見、匯出)
                                </span>
                            </td>
                        </tr>
                        <tr id="trFTDate12" runat="server">
                            <td class="bluecol" width="20%">結訓期間 </td>
                            <td class="whitecol" colspan="3">
                                <span id="span02" runat="server">
                                    <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
							    <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                    (配合列印其他意見、匯出)
                                </span>
                            </td>
                        </tr>
                        <tr id="trPackageType" runat="server">
                            <td class="bluecol" width="20%">包班種類 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PackageType" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="A" Selected="True">全部</asp:ListItem>
                                    <asp:ListItem Value="2">企業包班</asp:ListItem>
                                    <asp:ListItem Value="3">聯合企業包班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <%--<tr><td class="bluecol_need">調查表版本</td><td class="whitecol" colspan="3"><asp:DropDownList ID="ddlFACVER" runat="server">
                            <asp:ListItem Value="1" Selected="True">1</asp:ListItem><asp:ListItem Value="2">2</asp:ListItem></asp:DropDownList></td></tr>--%>
                        <tr>
                            <td class="bluecol" width="20%">交叉查詢選項</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="FuncID" runat="server"></asp:DropDownList>
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

                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table width="100%">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                <%--export--%>
                                <asp:Button ID="BtnPrint3" runat="server" Text="列印明細" Visible="False" CssClass="asp_Export_M"></asp:Button>
                                <%--export--%>
                                <asp:Button ID="btnPrint4" runat="server" Text="列印其他意見" Visible="False" CssClass="asp_Export_M"></asp:Button>
                                <%--export--%>
                                <asp:Button ID="btnExport5" runat="server" Text="匯出調查統計分析表" Visible="False" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
<%-- <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server"><tr><td>
    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False"><Columns>
    <asp:TemplateColumn HeaderText="調查項目"></asp:TemplateColumn><asp:TemplateColumn HeaderText="調查內容"></asp:TemplateColumn>
    <asp:TemplateColumn HeaderText="一般身分者"></asp:TemplateColumn></Columns></asp:DataGrid></td></tr></table>--%>
        <asp:HiddenField ID="Hid_LID" runat="server" />
        <asp:HiddenField ID="Hid_YEARS" runat="server" />
    </form>
</body>
</html>
