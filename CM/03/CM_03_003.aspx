<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_003.aspx.vb" Inherits="WDAIIP.CM_03_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練人數綜合查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }

        function GetOrg() {
            var msg = '';
            var DistID = getRadioValue(document.getElementsByName('DistID'));
            var TPlanID = getRadioValue(document.getElementsByName('TPlanID'));
            if (DistID == '') msg += '請先選擇轄區\n';
            if (TPlanID == '') msg += '請先選擇訓練計畫\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
            wopen('../../Common/MainOrg.aspx?DistID=' + DistID + '&TPlanID=' + TPlanID + '&BtnName=Button3', '查詢機構', 400, 400, 1);
        }

        function ClearData() {
            document.getElementById('PlanID').value = '';
            document.getElementById('center').value = '';
            document.getElementById('RIDValue').value = '';
            for (var i = document.form1.OCID.options.length - 1; i >= 0; i--) {
                document.form1.OCID.options[i] = null;
            }
            document.getElementById('OCID').style.display = 'none';
            document.getElementById('msg').innerHTML = '請先選擇機構';

        }

        function ReStart() {
            window.scroll(0, document.body.scrollHeight);
        }

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
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;綜合動態報表</asp:Label>
                    <%-- 首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;訓練人數綜合查詢--%>
                </td>
            </tr>
        </table>

        <%--<table id="FrameTable2" border="0" cellspacing="1" cellpadding="1" width="100%">
        </table>--%>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td class="bluecol" width="20%">動態報表 </td>
                <td class="whitecol" width="80%">
                    <uc1:WUC2 runat="server" ID="WUC2" />
                </td>
            </tr>
            <tr id="trSYEAR" runat="server">
                <td class="bluecol" width="20%">年度
                </td>
                <td class="whitecol">
                    <asp:DropDownList ID="ddlSYEAR" runat="server">
                    </asp:DropDownList>
                </td>
            </tr>
            <tr>
                <td class="bluecol">開訓區間
                </td>
                <td class="whitecol" runat="server">
                    <asp:TextBox ID="STDate1" runat="server" Columns="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    ~<asp:TextBox ID="STDate2" runat="server" Columns="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('STDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">結訓區間
                </td>
                <td class="whitecol" runat="server">
                    <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    ~<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">轄區
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                    </asp:CheckBoxList>
                    <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練計畫
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="0" CellPadding="0" RepeatColumns="3">
                    </asp:CheckBoxList>
                    <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">縣市
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="CityID" runat="server" RepeatColumns="8" RepeatDirection="Horizontal" CssClass="font">
                    </asp:CheckBoxList>
                    <input id="CityHidden" type="hidden" value="0" name="CityHidden" runat="server">
                </td>
            </tr>
            <tr>
                <td class="bluecol">訓練機構
                </td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                    <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    <input id="PlanID" type="hidden" name="PlanID" runat="server">
                    <asp:Button ID="Button3" runat="server" Text="查詢班級" CssClass="asp_Export_M"></asp:Button>(勾選班級後會省略開訓區間、結訓區期的條件)
                </td>
            </tr>
            <tr>
                <td class="bluecol">班別
                </td>
                <td class="whitecol">
                    <asp:ListBox ID="OCID" runat="server" Rows="6" SelectionMode="Multiple" Width="40%"></asp:ListBox>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>(按Ctrl可以複選班級)
                </td>
            </tr>
            <tr>
                <td class="bluecol">預算來源
                </td>
                <td class="whitecol">
                    <asp:CheckBoxList ID="BudgetList" runat="server" RepeatDirection="Horizontal" CssClass="font">
                    </asp:CheckBoxList>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need">統計項目</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblMode" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                    </asp:RadioButtonList>
                    <%-- <asp:ListItem Value="0" Selected="True">身分別</asp:ListItem>
                        <asp:ListItem Value="1">年齡</asp:ListItem>
                        <asp:ListItem Value="2">訓練職類</asp:ListItem>
                        <asp:ListItem Value="3">教育程度</asp:ListItem>
                        <asp:ListItem Value="4">性別</asp:ListItem>
                        <asp:ListItem Value="5">通俗職類</asp:ListItem>
                        <asp:ListItem Value="6">訓練業別</asp:ListItem>
                        <asp:ListItem Value="7">縣市別</asp:ListItem>
                        <asp:ListItem Value="9">上課時數</asp:ListItem>--%>
                    <%--<asp:ListItem Value="8">失業週數</asp:ListItem>--%>
                    <%-- <br> <asp:Label ID="Label1" runat="server" ForeColor="Red">(選擇身分別統計項目時,因學習券情況特殊,請勿與其他計畫勾選查詢,即學習券單選)</asp:Label>--%>
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
            <tr>
                <td colspan="2" align="center">
                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_Export_M"></asp:Button>&nbsp;
					<%--<asp:Button ID="Print" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>--%>
                </td>
            </tr>
        </table>

        <div id="Div1" runat="server">
            <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" ShowFooter="True" AutoGenerateColumns="False">
                <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                <HeaderStyle CssClass="head_navy"></HeaderStyle>
                <FooterStyle HorizontalAlign="Center"></FooterStyle>
                <Columns>
                    <asp:BoundColumn DataField="Title" HeaderText="統計項目" FooterText="合計">
                        <HeaderStyle Width="20%"></HeaderStyle>
                        <FooterStyle HorizontalAlign="Left"></FooterStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="JoinStudent" HeaderText="開訓人數">
                        <HeaderStyle Width="20%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        <FooterStyle HorizontalAlign="Center"></FooterStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="FinStudent" HeaderText="結訓人數">
                        <HeaderStyle Width="20%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        <FooterStyle HorizontalAlign="Center"></FooterStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ReStudent" HeaderText="離退訓人數">
                        <HeaderStyle Width="20%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                        <FooterStyle HorizontalAlign="Center"></FooterStyle>
                    </asp:BoundColumn>
                </Columns>
            </asp:DataGrid>
        </div>

        <p align="center">
            <asp:Button ID="Button4" runat="server" Text="回上一頁" CssClass="asp_Export_M"></asp:Button>&nbsp;
			<asp:Button ID="btnExport1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
        </p>

    </form>
</body>
</html>
