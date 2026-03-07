<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_020.aspx.vb" Inherits="WDAIIP.SD_14_020" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>材料明細表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
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
        function getrblvalue1(rbln) {
            var val = "";
            //得到radiobuttonlist
            var vRbtid = document.getElementById(rbln);
            //得到所有radio
            var vRbtidList = vRbtid.getElementsByTagName("INPUT");
            for (var i = 0; i < vRbtidList.length; i++) {
                if (vRbtidList[i].checked) {
                    //var text =vRbtid.cells[i].innerText;
                    val = vRbtidList[i].value;
                    return val;
                    //alert("選中項的text值為" text ",value值為" value);
                }
            }
            return val;
        }
        /*
        function CheckPrint() {
            var v_rblCT1 = getrblvalue1("rblClassType1");
            alert('v_rblCT1:' + v_rblCT1); return false;
            if (v_rblCT1 != "0") {
                if (document.getElementById('OCIDValue').value == '') { alert('請選擇班級'); return false; }
            }
            else {
                if (document.getElementById('PlanIDValue').value == '') { alert('請選擇班級'); return false; }
            }
        }
        */

        //checkbox 全選
        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                if (MyTable.rows[i].cells[0].children[0]) {
                    MyTable.rows[i].cells[0].children[0].checked = Flag;
                }
                //SelectOCID(Flag,MyTable.rows[i].cells[0].children[0].value);
            }
        }

        //SD_14_015
        //清除		
        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
        }

        //選班
        function choose_class() {
            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';
            openClass('../02/SD_02_ch.aspx?&RID=' + document.getElementById('RIDValue').value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <table id="Table1" class="font" border="0" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;<FONT color="#990000">材料明細表</FONT></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" width="100%">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" name="Hidden2" type="hidden" runat="server">
                                <input id="Button2" name="Button2" value="..." type="button" runat="server" class="button_b_Mini">
                                <span style="position: absolute; display: none" id="HistoryList2">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">班級狀態</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="rblClassType1" runat="server" AutoPostBack="True" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="0">未轉班</asp:ListItem>
                                    <asp:ListItem Value="1">已轉班</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="ClassTR" runat="server">
                            <td class="bluecol">職類/班別</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" value="..." type="button" class="button_b_Mini">
                                <input id="Button4" name="Button4" value="清除" type="button" runat="server" class="asp_button_S">
                                <input id="TMIDValue1" name="TMIDValue1" type="hidden" runat="server">
                                <input id="OCIDValue1" name="OCIDValue1" type="hidden" runat="server">
                                <span style="position: absolute; left: 270px; display: none" id="HistoryList">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="tr_AppStage_TP28" runat="server">
                            <td class="bluecol">申請階段</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="AppStage2" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font"></asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="TRPlanPoint28" runat="server">
                            <td class="bluecol">計畫</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="PlanPoint" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow" CssClass="font">
                                    <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
                                    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="4" align="center" class="whitecol">
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4" align="center">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></td>
                        </tr>
                    </table>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDate" HeaderText="開訓日期">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDate" HeaderText="結訓日期">
                                            <HeaderStyle Width="100px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderStyle-Width="7%">
                                            <%--<input type="checkbox" onclick="SelectAll(this.checked);">--%>
                                            <HeaderTemplate>功能</HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                            <ItemTemplate>
                                                <asp:Button ID="btnPrint1" runat="server" class="asp_Export_M" Text="列印" CommandName="Print1"></asp:Button>
                                                <%--<input id="PrintRpt1" type="button" value="列印" runat="server" class="asp_Export_M" />--%>
                                                <%--<input type="checkbox" id="chkbox1" runat="server" name="chkbox1">--%>
                                                <input id="hidOCID" type="hidden" runat="server" name="hidOCID">
                                                <input id="hidPlanID" type="hidden" runat="server" name="hidPlanID">
                                                <input id="hidComIDNO" type="hidden" runat="server" name="hidComIDNO">
                                                <input id="hidSeqNo" type="hidden" runat="server" name="hidSeqNo">
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
                        <%--<tr>
                            <td align="center">
                                <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button></td>
                        </tr>--%>
                    </table>
                </td>
            </tr>
        </table>
        <input id="Years" name="Years" type="hidden" runat="server" />
        <input id="orgid" name="orgid" type="hidden" runat="server" />
        <input id="OCIDValue" name="OCIDValue" type="hidden" runat="server" />
        <input id="PLANIDValue" name="PLANIDValue" type="hidden" runat="server" />
        <input id="ComIDNOValue" name="ComIDNOValue" type="hidden" runat="server" />
        <input id="SeqNoValue" name="SeqNoValue" type="hidden" runat="server" />
        <input id="PCSValue" name="PCSValue" type="hidden" runat="server" />
    </form>
</body>
</html>
