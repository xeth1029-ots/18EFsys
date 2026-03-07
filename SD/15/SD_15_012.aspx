<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_15_012.aspx.vb" Inherits="WDAIIP.SD_15_012" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>綜合動態報表-綜合查詢統計表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script language="javascript" type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/date-picker.js"></script>
    <script language="javascript" type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script language="javascript" type="text/javascript">
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個
            //alert(getCheckBoxListValue(obj));
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0); //記憶
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                //若有全選
                if (getCheckBoxListValue(obj).charAt(0) == '1') {
                    myallcheck.checked = false; //全選改為false
                    document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0); //記憶
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;統計表&gt;&gt;綜合動態報表</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_sch" runat="server" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">動態報表 </td>
                            <td class="whitecol" width="80%">
                                <uc1:WUC1 runat="server" ID="WUC1" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">年度</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="True">
                                </asp:DropDownList>
                                <%--<asp:RequiredFieldValidator ID="MustYear" runat="server" ErrorMessage="請選擇年度" Display="Dynamic" ControlToValidate="yearlist" CssClass="font"></asp:RequiredFieldValidator>--%>
                            </td>
                        </tr>
                        <tr id="trPlanKind" runat="server">
                            <td class="bluecol">計畫別</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="OrgPlanKind" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">轄區</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cblDistid" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">申請階段 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cbl_AppStage" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                                    <asp:ListItem Value="1">1:上半年</asp:ListItem>
                                    <asp:ListItem Value="2">2:下半年</asp:ListItem>
                                    <asp:ListItem Value="3">3:政策性產業</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="70%" onfocus="this.blur()"></asp:TextBox>
                                <input id="Org" value="..." type="button" name="Org" runat="server" class="button_b_Mini" />
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="ComidValue" type="hidden" name="ComidValue" runat="server" />
                            </td>
                        </tr>
                        <%--<tr id="Org_TR" runat="server"><td class="bluecol">機構</td><td class="whitecol"><asp:TextBox ID="center" runat="server" Width="410px"></asp:TextBox>
<input id="BtnOpenOrg1" value="..." type="button" runat="server" class="button_b_Mini"  /><input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
<input id="PlanID" type="hidden" name="PlanID" runat="server" /></td></tr>--%>
                        <tr id="trPackageType" runat="server">
                            <td class="bluecol">包班種類
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="PackageType" runat="server" RepeatDirection="Horizontal" RepeatColumns="7" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="PackageHidden" type="hidden" value="0" name="PackageHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">辦訓地縣市
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="Tcitycode" runat="server" RepeatDirection="Horizontal" RepeatColumns="7" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="TcityHidden" type="hidden" value="0" name="TcityHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">立案地縣市
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="Ocitycode" runat="server" RepeatDirection="Horizontal" RepeatColumns="7" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="OcityHidden" type="hidden" value="0" name="OcityHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練業別
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="GovClassName" runat="server" RepeatDirection="Horizontal" RepeatColumns="9" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="GovClassHidden" type="hidden" value="0" name="GovClassHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練職能
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="CCID" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="CCIDHidden" type="hidden" value="0" name="CCIDHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">課程分類
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="cblDepot12" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="HidcblDepot12" type="hidden" value="0" runat="server">
                            </td>
                        </tr>
                        <tr id="KID_6_TR" runat="server">
                            <td class="bluecol">新興產業
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="KID_6" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="KID_6_hid" type="hidden" value="0" name="HID_DepID_6" runat="server">
                            </td>
                        </tr>
                        <tr id="KID_10_TR" runat="server">
                            <td class="bluecol">重點服務業
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="KID_10" runat="server" RepeatDirection="Horizontal" RepeatColumns="4" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="KID_10_hid" type="hidden" value="0" name="KID_10_hid" runat="server">
                            </td>
                        </tr>
                        <tr id="KID_4_TR" runat="server">
                            <td class="bluecol">新興智慧型產業
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="KID_4" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="KID_4_hid" type="hidden" value="0" name="KID_4_hid" runat="server">
                            </td>
                        </tr>
                        <%--<tr id="KID_17_tr" runat="server"><td class="bluecol">政府政策性產業<br />(108年之後不使用此欄)</td>
<td class="whitecol"><asp:CheckBoxList ID="KID_17" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" CssClass="whitecol">
</asp:CheckBoxList><input id="KID_17_hid" type="hidden" value="0" name="KID_17_hid" runat="server">
<asp:CheckBoxList ID="KID_19" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" CssClass="whitecol">
</asp:CheckBoxList><input id="KID_19_hid" type="hidden" value="0" runat="server" /></td></tr>--%>
                        <tr id="KID_20_tr" runat="server">
                            <td class="bluecol">政府政策性產業<br />
                                (114年前(不含))</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="KID_20" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" CssClass="whitecol"></asp:CheckBoxList>
                                <input id="KID_20_hid" type="hidden" value="0" name="KID_20_hid" runat="server" /></td>
                        </tr>
                        <tr id="KID_25_tr" runat="server">
                            <td class="bluecol">政府政策性產業<br />
                                (114年後(含))</td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="KID_25" runat="server" RepeatDirection="Horizontal" RepeatColumns="5" CssClass="whitecol"></asp:CheckBoxList>
                                <input id="KID_25_hid" type="hidden" value="0" name="KID_25_hid" runat="server" /></td>
                        </tr>
                        <tr>
                            <td class="bluecol">是否為學分班
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="PointYN" runat="server" RepeatDirection="Horizontal" CssClass="whitecol">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">是否核定
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Apppass" runat="server" RepeatDirection="Horizontal" CssClass="whitecol">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">是否結訓
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Endclass" runat="server" RepeatDirection="Horizontal" CssClass="whitecol">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">是否撥款
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Appmoney" runat="server" RepeatDirection="Horizontal" CssClass="whitecol">
                                    <asp:ListItem Value="A">不區分</asp:ListItem>
                                    <asp:ListItem Value="1">是</asp:ListItem>
                                    <asp:ListItem Value="0">否</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">是否停辦
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="Stopclass" runat="server" RepeatDirection="Horizontal" CssClass="whitecol">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">&nbsp;檢送研提資料</td>
                            <td class="whitecol"><%--檢送資料-未檢送-含未檢送研提資料--%>
                                <asp:CheckBox ID="CB_DataNotSent_SCH" runat="server" Text="含未檢送研提資料" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓日期
                            </td>
                            <td class="whitecol" runat="server">
                                <asp:TextBox ID="SDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>&nbsp;<img alt="" style="cursor: pointer" onclick="javascript:show_calendar('<%= SDate1.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                &nbsp;~&nbsp;<asp:TextBox ID="SDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>
                                <img alt="" style="cursor: pointer" onclick="javascript:show_calendar('<%= SDate2.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">結訓日期
                            </td>
                            <td class="whitecol" runat="server">
                                <asp:TextBox ID="EDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>&nbsp;<img alt="" style="cursor: pointer" onclick="Javascript:show_calendar('<%= EDate1.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                &nbsp;~&nbsp;<asp:TextBox ID="EDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>&nbsp;<img alt="" style="cursor: pointer" onclick="Javascript:show_calendar('<%= EDate2.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">撥款日期
                            </td>
                            <td class="whitecol" runat="server">
                                <asp:TextBox ID="AllotDate1" runat="server" Columns="10" MaxLength="10"></asp:TextBox>&nbsp;<img alt="" style="cursor: pointer" onclick="javascript:show_calendar('<%= AllotDate1.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                &nbsp;~&nbsp;<asp:TextBox ID="AllotDate2" runat="server" Columns="10" MaxLength="10"></asp:TextBox>&nbsp;<img alt="" style="cursor: pointer" onclick="javascript:show_calendar('<%= AllotDate2.ClientId %>','','','CY/MM/DD');" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">匯出欄位
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="ChbExit" runat="server" RepeatDirection="Horizontal" RepeatColumns="3" CssClass="whitecol">
                                </asp:CheckBoxList>
                                <input id="ChbExitHidden" type="hidden" value="0" name="ChbExitHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">參數資料
                            </td>
                            <td class="whitecol">
                                <asp:CheckBox ID="Cbl_His_MV_DATA" runat="server" Text="使用歷史資料（為當天凌晨定版資料，匯出時間會較短）" />
                            </td>
                        </tr>
                        <%-- <tr><td class="bluecol">匯出檔案編碼格式</td><td class="whitecol"><asp:RadioButtonList ID="RBL_CharsetType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
<asp:ListItem Value="BIG5" Selected="True">BIG5</asp:ListItem><asp:ListItem Value="UTF8">UTF-8</asp:ListItem></asp:RadioButtonList></td></tr>--%>
                        <tr>
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                    <asp:ListItem Value="XLSX">XLSX</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="BtnExp" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    </p>
                    <p>
                        <font color="#ff0000">匯出欄位說明:<br />
                            1. 實際開訓人次：開訓後14天實際錄訓人數<br />
                            2. 結訓人次：排除不開班及離退訓，學員資料確認，班級結訓，結訓成績登錄功能有選擇是否有取得學分資格<br />
                            3. 撥款人次：排除不開班及離退訓，學員資料確認，班級結訓，結訓成績登錄功能有選擇是否有取得學分資格，補助撥款功能為通過<br />
                            4. [單一人時成本]：固定費用÷人數÷時數<br />
                            &nbsp;&nbsp;&nbsp; [材料費人時成本]：單一人時成本 * 各職類材料費編列比率上限<br />
                        </font>
                    </p>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="hid_ssYears" runat="server" />
        <asp:HiddenField ID="Hid_GovClassT" runat="server" />
    </form>
</body>
</html>
