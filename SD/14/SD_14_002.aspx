<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_002.aspx.vb" Inherits="WDAIIP.SD_14_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練班別計畫表 / 開班計畫總表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <%--<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function rblFONTTYPE_CHG1() {
            $("#DataGridTable").hide();
        }
    </script>
    <script type="text/javascript" language="javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181023
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);

        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
        }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';
            openClass('../02/SD_02_ch.aspx?&RID=' + RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;表單列印&gt;&gt;訓練班別計畫表</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級狀態 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="Radio1" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" AutoPostBack="True">
                                    <asp:ListItem Value="0">未轉班</asp:ListItem>
                                    <asp:ListItem Value="1">已轉班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="ClassTR" runat="server">
                            <td class="bluecol">職類/班別 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="Button4" type="button" value="清除" name="Button4" runat="server" class="asp_button_M">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">開訓期間 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>～
                                <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr id="TR_2" runat="server">
                            <td class="bluecol">課程審核狀況 </td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="rdlResult" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="" Selected="True">不拘</asp:ListItem>
                                    <asp:ListItem Value="Y">通過</asp:ListItem>
                                    <asp:ListItem Value="N">不通過</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="TR_3" runat="server">
                            <td class="bluecol">班級名稱 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ClassName" runat="server" Width="40%" MaxLength="100"></asp:TextBox></td>
                        </tr>
                        <tr id="TR_4" runat="server">
                            <td class="bluecol">課程關鍵字 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ClassKW1" runat="server" Width="40%" MaxLength="100"></asp:TextBox></td>
                        </tr>
                        <tr id="TR_5" runat="server">
                            <td class="bluecol">課程代碼 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ClassKW2" runat="server" Width="30%" MaxLength="30"></asp:TextBox></td>
                        </tr>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="AppStage" runat="server"></asp:DropDownList></td>
                        </tr>
                        <tr id="TR_rblFONTTYPE" runat="server">
                            <td class="bluecol">列印字型選擇</td>
                            <td class="whitecol" colspan="3">
                                <%--'1:細明體/2:標楷體(def)--%>
                                <asp:RadioButtonList ID="rblFONTTYPE" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="1">細明體</asp:ListItem>
                                    <asp:ListItem Value="2" Selected="True">標楷體</asp:ListItem>
                                </asp:RadioButtonList></td>
                        </tr>
                        <tr id="tr_RBListExpType" runat="server">
                            <td class="bluecol">匯出檔案格式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                    <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </div>
                    <div align="center">
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle Width="30%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                            <HeaderStyle Width="36%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn>
                                            <HeaderStyle Width="6%" />
                                            <ItemTemplate>
                                                <input id="OCID" type="hidden" runat="server" />
                                                <input id="PlanID" type="hidden" runat="server" />
                                                <input id="ComIDNO" type="hidden" runat="server" />
                                                <input id="SeqNo" type="hidden" runat="server" />
                                                <input id="PrintRpt1" type="button" value="列印" runat="server" class="asp_Export_M" />
                                                <%--<input id="PrintRpt2" type="button" value="列印" runat="server" />--%>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Button ID="BTN_EXPORT1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>

                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_Radio1" runat="server" />
        <%--Hid_Radio1::'0:未轉班,1:已轉班--%>
    </form>
</body>
</html>
