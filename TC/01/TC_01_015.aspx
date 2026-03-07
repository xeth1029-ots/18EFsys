<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_015.aspx.vb" Inherits="WDAIIP.TC_01_015" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練單位處分功能</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <%--<script type="text/javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181018
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);       

        function IsDate(MyDate) {
            if (MyDate != '') {
                if (!checkDate(MyDate))
                    return false;
            }
            return true;
        }

        function chkdata() {
            var msg = '';
            //debugger;
            if (document.form1.txt_ComIDNO.value != '') {
                if (document.form1.txt_ComIDNO.value.length > 8) {
                    msg += '統一編號長度超過範圍\n';
                }
            }
            else {
                msg += '請輸入統一編號\n';
            }
            if (document.form1.txt_No.value == '') {
                msg += '處分文號為必填\n';
            }
            if (isEmpty(document.form1.ddlOBTERMS)) {
                msg += '請選擇處分緣由\n';
            }
            if (document.form1.txt_OBSdate.value != '') {
                if (!IsDate(document.form1.txt_OBSdate.value)) {
                    msg += '處分日期年月日有誤';
                }
            }
            else {
                msg += '處分日期年月日為必填\n';
            }
            if (isEmpty(document.form1.ddl_OBYears)) {
                msg += '請選擇處分年限\n';
            }
            if (document.form1.txt_OBComment.value == '') {
                msg += '處分事由為必填\n';
            }
            else {
                if (checkMaxLen(document.form1.txt_OBComment.value, 300 * 2)) {
                    msg += '處分事由長度不可超過300字元\n';
                }
            }
            if (msg != '') {
                window.alert(msg);
                return false;
            }
            else {
                msg = '';
                msg += '\n請確認資料是否無誤,儲存後資料將不可修改\n\n';
                msg += '如確認資料無誤後,請按下確定,謝謝!!\n';
                return confirm(msg);
            }
        }

        //20180706 add 班級名稱
        function choose_class() {
            var RIDX1 = '0';
            if (document.form1.RIDValueX1) {
                if (document.form1.RIDValueX1.value.length > 0)
                    RIDX1 = document.form1.RIDValueX1.value;
            }
            if (document.form1.TMID1)
                document.form1.TMID1.value = '';
            if (document.form1.TMIDValue1)
                document.form1.TMIDValue1.value = '';
            if (document.form1.OCID1)
                document.form1.OCID1.value = '';
            if (document.form1.OCIDValue1)
                document.form1.OCIDValue1.value = '';
            if (document.form1.hidLockTime1)
                document.form1.hidLockTime1.value = '1';  //1:鎖定
            //openClass('../02/SD_02_ch.aspx?special=11&RID=' + document.form1.RIDValueX1.value);
            openClass('../../SD/02/SD_02_ch.aspx?special=11&RID=' + RIDX1);
        }

        function clearDate(objId) {
            var myObj = document.getElementById(objId);
            myObj.value = "";
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td align="center">
                    <table class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;訓練單位處分功能</asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <asp:Panel ID="Panel1" runat="server" Visible="True">
                        <table id="Table3" class="table_sch" cellpadding="1" cellspacing="1">
                            <tr>
                                <td class="bluecol" width="20%">計畫別</td>
                                <td class="whitecol" colspan="3">
                                    <asp:DropDownList ID="ddlTPlanIDSch" runat="server"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">原處分分署</td>
                                <td class="whitecol" width="30%">
                                    <asp:DropDownList ID="DistID" runat="server"></asp:DropDownList></td>
                                <td class="bluecol" width="20%">處分年度</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="Years" runat="server"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練機構</td>
                                <td colspan="3" class="whitecol">
                                    <asp:TextBox ID="center" runat="server" Width="45%" onfocus="this.blur()"></asp:TextBox>
                                    <input id="Org" value="..." type="button" name="Org" runat="server" class="button_b_Mini">
                                    <input id="RIDValue" size="3" type="hidden" name="RIDValue" runat="server">
                                    <input id="orgid_value" type="hidden" name="orgid_value" runat="server">
                                    <span style="position: absolute; display: none" id="HistoryList2">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">統一編號</td>
                                <td class="whitecol" width="30%">
                                    <asp:TextBox ID="ComidValue" runat="server" Width="80%" MaxLength="10"></asp:TextBox></td>
                                <td class="bluecol">機構名稱</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtOrgName" runat="server" MaxLength="30" Columns="30"></asp:TextBox></td>
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
                                <td class="whitecol" align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                    <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnAdds" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                            </tr>
                        </table>
                        <table id="tb_Sch" border="0" cellspacing="0" cellpadding="0" width="100%" runat="server">
                            <tr>
                                <td align="center">
                                    <div id="Div1" runat="server">
                                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <Columns>
                                                <asp:BoundColumn HeaderText="序號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="PlanName" HeaderText="計畫別">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="DistName" HeaderText="原處分分署">
                                                    <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱">
                                                    <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="ComIDNO" HeaderText="統一編號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="處分緣由">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="labOBTERMS" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn DataField="OBSDATE" HeaderText="處分起日">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="OBYears" HeaderText="年限">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="C_PunishPeriod" HeaderText="處分期間">
                                                    <HeaderStyle HorizontalAlign="Center" Width="9%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="OBComment" HeaderText="事由">
                                                    <HeaderStyle HorizontalAlign="Center" Width="9%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="OBComment" HeaderText="事由">
                                                    <HeaderStyle HorizontalAlign="Center" Width="9%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="功能">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:LinkButton ID="lbtView" runat="server" Text="檢視" CommandName="view" CssClass="linkbutton"></asp:LinkButton>
                                                        <asp:LinkButton ID="lbtEdit" runat="server" Text="修改" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                                                        <asp:LinkButton ID="lbtDel" runat="server" Text="刪除" CommandName="del" CssClass="linkbutton"></asp:LinkButton>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                            </Columns>
                                            <PagerStyle Visible="False"></PagerStyle>
                                        </asp:DataGrid>
                                    </div>
                                    <div>
                                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    </div>
                                    <div id="DivOutputDoc" runat="server" visible="false">
                                        <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="False" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                            <Columns>
                                                <asp:BoundColumn HeaderText="序號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="2%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="DistName" HeaderText="原處分分署">
                                                    <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="PlanName" HeaderText="計畫別">
                                                    <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="C_YEAR" HeaderText="課程所屬年度">
                                                    <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="COMIDNO" HeaderText="訓練單位統一編號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="OrgName" HeaderText="單位名稱">
                                                    <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="C_NAME" HeaderText="課程名稱">
                                                    <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="C_PERIOD" HeaderText="訓練期間">
                                                    <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_APPLYPRICE" HeaderText="申請金額">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_AUTHPRICE" HeaderText="核定金額">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="OBComment" HeaderText="異常事由(請詳述)">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="OBSDATE" HeaderText="處分日期">
                                                    <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="OBNUM" HeaderText="處分文號">
                                                    <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="OBFACT" HeaderText="處分事實">
                                                    <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:TemplateColumn HeaderText="處分依據">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="labOBTERMS" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn DataField="MY_OBYEARS" HeaderText="停權期限">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="C_PunishPeriod" HeaderText="處分期間">
                                                    <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_LAW1" HeaderText="是否會辦政風">
                                                    <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_LAW2" HeaderText="是否移送檢調">
                                                    <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_TRANSFER" HeaderText="移送情形">
                                                    <HeaderStyle HorizontalAlign="Center" Width="3%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_JUDGE_1" HeaderText="檢調偵查/判決情形 (判決日期、文號)">
                                                    <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_JUDGE_2" HeaderText="檢調偵查/判決情形 (判決事實)">
                                                    <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_JUDGE_3" HeaderText="後續待辦事項 (追繳款項、強制執行狀況等)">
                                                    <HeaderStyle HorizontalAlign="Center" Width="6%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_NOTE" HeaderText="備註">
                                                    <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                                <asp:BoundColumn DataField="MY_MODIFYDATE" HeaderText="異動日期">
                                                    <HeaderStyle HorizontalAlign="Center" Width="4%"></HeaderStyle>
                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                </asp:BoundColumn>
                                            </Columns>
                                            <PagerStyle Visible="False"></PagerStyle>
                                        </asp:DataGrid>
                                    </div>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="Panel2" runat="server" Visible="True">
                        <table class="table_sch" cellpadding="1" cellspacing="1">
                            <tr>
                                <td class="bluecol" width="20%">計畫別</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddlTPlanID" runat="server"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol">原處分分署</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddl_DistID" runat="server"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol">機構名稱</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="centerX1" runat="server" Width="45%" onfocus="this.blur()"></asp:TextBox>
                                    <input id="OrgX1" value="..." type="button" name="OrgX1" runat="server" class="asp_button_Mini" />
                                    <input id="RIDValueX1" size="3" type="hidden" name="RIDValueX1" runat="server" />
                                    <input id="orgid_valueX1" type="hidden" name="orgid_valueX1" runat="server" />
                                    <span style="position: absolute; display: none" id="HistoryList2X1">
                                        <asp:Table ID="HistoryRIDX1" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">統一編號</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_ComIDNO" runat="server" Width="30%" MaxLength="10" onkeypress="return event.charCode >= 48 && event.charCode <= 57"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級名稱</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="35%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="45%"></asp:TextBox>
                                    <input id="Button5" onclick="choose_class();" value="..." type="button" runat="server" class="asp_button_Mini" />
                                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                    <span style="position: absolute; display: none; left: 35%" id="HistoryList">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">申請金額</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_ApplyPrice" runat="server" Width="20%" MaxLength="6" onkeypress="return event.charCode >= 48 && event.charCode <= 57"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">核定金額</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_AuthPrice" runat="server" Width="20%" MaxLength="6" onkeypress="return event.charCode >= 48 && event.charCode <= 57"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分文號</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_No" runat="server" Width="30%" MaxLength="30"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分緣由 </td>
                                <td style="height: 4%" class="whitecol">
                                    <asp:DropDownList ID="ddlOBTERMS" runat="server"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分日期</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_OBSdate" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox>&nbsp;
                                    <span id="span1" runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txt_OBSdate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分年限</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="ddl_OBYears" runat="server">
                                        <asp:ListItem Value="0">0年</asp:ListItem>
                                        <asp:ListItem Value="1">1年</asp:ListItem>
                                        <asp:ListItem Value="2">2年</asp:ListItem>
                                        <asp:ListItem Value="3">3年</asp:ListItem>
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分期間</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_PunishPeriod" runat="server" Width="50%" onfocus="this.blur()"></asp:TextBox>&nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">處分事由</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_OBComment" runat="server" Width="50%" MaxLength="150" TextMode="MultiLine" Rows="5"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分事實</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_OBFact" runat="server" Width="50%" MaxLength="150" TextMode="MultiLine" Rows="5"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">是否會辦政風</td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rbl_IsLaw1" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="Y">是</asp:ListItem>
                                        <asp:ListItem Value="N" Selected="True">否</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">是否移送檢調</td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rbl_IsLaw2" runat="server" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="Y">是</asp:ListItem>
                                        <asp:ListItem Value="N" Selected="True">否</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">移送情形</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_Transfer" runat="server" Width="50%" MaxLength="150" TextMode="MultiLine" Rows="5"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">檢調偵查/判決日期</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_JudgeDate" runat="server" Width="15%" onfocus="this.blur()"></asp:TextBox>&nbsp;
                                    <span id="span2" runat="server">
                                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txt_JudgeDate.ClientId %>','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="30" height="30" /></span>
                                    <span id="span3" runat="server">
                                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= txt_JudgeDate.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">檢調偵查/判決文號</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_JudgeNum" runat="server" Width="40%" MaxLength="30"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">檢調偵查/判決事實</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_JudgeFact" runat="server" Width="50%" MaxLength="150" TextMode="MultiLine" Rows="5"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">後續待辦事項<br />
                                    (追繳款項、強制執行狀況等)</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_Tudo" runat="server" Width="50%" MaxLength="150" TextMode="MultiLine" Rows="5"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">備註</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_Note" runat="server" Width="50%" MaxLength="150" TextMode="MultiLine" Rows="5"></asp:TextBox>

                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">異動日期</td>
                                <td class="whitecol">
                                    <asp:Label ID="labModifyDate" runat="server" Text=""></asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btn_Save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
                                    <asp:Button ID="btn_lev" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <asp:Panel ID="Panel3" runat="server" Visible="True">
                        <table class="table_sch" cellpadding="1" cellspacing="1">
                            <tr>
                                <td class="bluecol" width="20%">計畫別</td>
                                <td class="whitecol" width="80%">
                                    <asp:Label ID="lbl_PlanName" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">原處分分署</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_DistID" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">機構名稱</td>
                                <td class="whitecol">
                                    <asp:Label ID="lab_OrgName" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">廠商統一編號</td>
                                <td class="whitecol">
                                    <asp:Label ID="lab_ComIDNO" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級名稱</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_CRName" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">申請金額</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_ApplyPrice" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">核定金額</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_AuthPrice" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分文號</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_No" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分緣由</td>
                                <td class="whitecol">
                                    <asp:Label ID="lblOBTERMS" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分日期</td>
                                <td class="whitecol">
                                    <asp:Label ID="lab_OBSDate" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分年限</td>
                                <td class="whitecol">
                                    <asp:Label ID="lab_OBYears" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分期間</td>
                                <td class="whitecol">
                                    <asp:Label ID="lab_PunishPeriod" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分事由</td>
                                <td class="whitecol">
                                    <asp:Label ID="lab_OBComment" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">處分事實</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_OBFact" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">是否會辦政風</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_IsLaw1" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">是否移送檢調</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_IsLaw2" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">移送情形</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_Transfer" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">檢調偵查/判決日期</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_JudgeDate" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">檢調偵查/判決文號</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_JudgeNum" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">檢調偵查/判決事實</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_JudgeFact" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">後續待辦事項<br />
                                    (追繳款項、強制執行狀況等)</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_Todo" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">備註</td>
                                <td class="whitecol">
                                    <asp:Label ID="lbl_Note" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">登錄系統者</td>
                                <td class="whitecol">
                                    <asp:Label ID="lab_accName" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">異動日期</td>
                                <td class="whitecol">
                                    <asp:Label ID="lab_Modifydate" runat="server" Text=""></asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="btn_lev2" runat="server" Text="離開" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
        </table>
        <input type="hidden" runat="server" id="hid_OBSN" />
    </form>
</body>
</html>
