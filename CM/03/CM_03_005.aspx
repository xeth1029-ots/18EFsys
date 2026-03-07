<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_005.aspx.vb" Inherits="WDAIIP.CM_03_005" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>生活津貼請領人數及請領金額查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">

        function GetOrg() {
            var msg = '';
            var DistID = getRadioValue(document.getElementsByName('DistID'));
            var TPlanID = getRadioValue(document.getElementsByName('TPlanID'));
            if (DistID == '') msg += '請先選擇轄區\n';
            if (TPlanID == '') msg += '請先選擇訓練計畫\n';
            if (msg != '') {
                alert(msg);
            }
            else {
                wopen('../../Common/MainOrg.aspx?DistID=' + DistID + '&TPlanID=' + TPlanID + '&BtnName=Button3', '查詢機構', 400, 400, 1);
            }
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


        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
            <tr>
                <td>
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;訓練需求管理&gt;&gt;統計分析&gt;&gt;<font color="#990000">生活津貼請領人數及請領金額查詢</font> </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table2" runat="server" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="80">結訓區間 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="FTDate1" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate1','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                                ~<asp:TextBox ID="FTDate2" runat="server" Columns="10"></asp:TextBox>
                                <span runat="server">
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('FTDate2','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol" style="height: 71px">轄區 </td>
                            <td class="whitecol" style="height: 71px">
                                <asp:CheckBoxList ID="DistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                </asp:CheckBoxList>
                                <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">訓練計畫 </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="TPlanID" runat="server" RepeatDirection="Horizontal" CssClass="font" CellSpacing="0" CellPadding="0" RepeatColumns="3">
                                </asp:CheckBoxList>
                                <input id="TPlanHidden" type="hidden" value="0" name="TPlanHidden" runat="server">
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Button5" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                    </p>
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" ShowFooter="True" AutoGenerateColumns="False" Width="100%">
                        <FooterStyle HorizontalAlign="Center"></FooterStyle>
                        <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="Title" HeaderText="統計項目" FooterText="合計">
                                <FooterStyle HorizontalAlign="Left"></FooterStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="SubsidyCount" HeaderText="請領人數">
                                <HeaderStyle Width="100px"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                <FooterStyle HorizontalAlign="Right"></FooterStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="SubsidySum" HeaderText="核發金額">
                                <HeaderStyle Width="120px"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                <FooterStyle HorizontalAlign="Right"></FooterStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="InJobCount" HeaderText="就業人數">
                                <HeaderStyle Width="100px"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                <FooterStyle HorizontalAlign="Right"></FooterStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn Visible="False" DataField="SubsidyJobCount" HeaderText="請領津貼就業人數">
                                <HeaderStyle Width="100px"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                <FooterStyle HorizontalAlign="Right"></FooterStyle>
                            </asp:BoundColumn>
                        </Columns>
                    </asp:DataGrid>
                    <tr>
                        <td align="center" colspan="2">
                            <%--<asp:button id="Button4" runat="server" Text="回上一頁"></asp:button>--%>
                        </td>
                    </tr>
            </tr>
        </table>
    </form>
</body>
</html>
