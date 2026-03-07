<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_05_015_R.aspx.vb" Inherits="WDAIIP.TR_05_015_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>TR_05_015_R</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }

        function search() {
            var msg = '';
            //if(document.form1.Syear.selectedIndex==0) msg+='請選擇年度\n';
            /*if(!isChecked(document.getElementsByName('TPlanID'))) msg+='請選擇訓練計畫\n';*/

            if (document.form1.STDate1.value != '') {
                if (!checkDate(document.form1.STDate1.value)) msg += '開訓日期的起始日不是正確的日期格式\n';
            }

            if (document.form1.STDate2.value != '') {
                if (!checkDate(document.form1.STDate2.value)) msg += '開訓日期的結束日不是正確的日期格式\n';
            }


            if (document.form1.FTDate1.value != '') {
                if (!checkDate(document.form1.FTDate1.value)) msg += '結訓日期的起始日不是正確的日期格式\n';
            }

            if (document.form1.FTDate2.value != '') {
                if (!checkDate(document.form1.FTDate2.value)) msg += '結訓日期的結束日不是正確的日期格式\n';
            }

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tbody>
                <tr>
                    <td>
                        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td>
                                    <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
										首頁&gt;&gt;訓練與就業需求管理&gt;&gt;統計分析&gt;&gt;參訓身分別
                                    </asp:Label>
                                </td>
                            </tr>
                        </table>
                        <table class="table_sch" id="Table2" cellspacing="1" cellpadding="1">
                            <tbody>
                                <tr>
                                    <td class="bluecol" style="width: 20%">年度
                                    </td>
                                    <td class="whitecol">
                                        <asp:DropDownList ID="Syear" runat="server">
                                        </asp:DropDownList>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">轄區
                                    </td>
                                    <td class="whitecol">
                                        <asp:CheckBoxList ID="DistID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
                                        </asp:CheckBoxList>
                                        <input id="DistHidden" type="hidden" value="0" runat="server" name="DistHidden" size="1">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">訓練計畫
                                    </td>
                                    <td class="whitecol">
                                        <asp:CheckBoxList ID="TPlanID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="3">
                                        </asp:CheckBoxList>
                                        <input id="TPlanHidden" type="hidden" value="0" runat="server" name="TPlanHidden" size="1">
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">開訓期間
                                    </td>
                                    <td class="whitecol">
                                        <asp:TextBox ID="STDate1" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
                                        <asp:TextBox ID="STDate2" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol" style="height: 6px">結訓期間
                                    </td>
                                    <td class="whitecol" style="height: 6px">
                                        <asp:TextBox ID="FTDate1" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"><font color="#000000">～
                                        <asp:TextBox ID="FTDate2" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></font>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol">匯出檔案格式</td>
                                    <td class="whitecol">
                                        <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                            <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                                            <asp:ListItem Value="ODS">ODS</asp:ListItem>
                                        </asp:RadioButtonList>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                        <p align="center" class="whitecol">
                            <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                            <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="btnExport" runat="server" Text="匯出Excel" CssClass="asp_Export_M"></asp:Button>
                        </p>
                        <p align="center">
                            <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                        </p>
                    </td>
                </tr>
            </tbody>
        </table>
        <table id="ResultTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <div id="Div1" runat="server">
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AllowPaging="True" CellPadding="8">
                            <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid>
                    </div>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
