<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TR_03_001.aspx.vb" Inherits="WDAIIP.TR_03_001" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>事業單位查詢</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function SelectAll() {
            var MyValue = getCheckBoxListValue('SCTID');
            var MyAllCheck = document.getElementById('SCTID_' + 0);

            if (document.getElementById('HidObj').value != MyValue.charAt(0)) {
                document.getElementById('HidObj').value = MyValue.charAt(0);
                for (var i = 1; i < MyValue.length; i++) {
                    var MyCheck = document.getElementById('SCTID_' + i);
                    MyCheck.checked = MyAllCheck.checked;
                }
            }
        }

        function HidTable() {
            if (document.getElementById('SearchTable')) {
                if (document.getElementById('SearchState').value == 1) {
                    document.getElementById('SearchTable').style.display = 'none';
                    document.getElementById('SearchState').value = 0;
                    document.getElementById('StateButton').innerHTML = '開啟查詢條件';
                }
                else {
                    document.getElementById('SearchTable').style.display = 'inline';
                    document.getElementById('SearchState').value = 1;
                    document.getElementById('StateButton').innerHTML = '關閉查詢條件';
                }
            }
        }

        function SetBDID(flag, BDIDValue) {
            if (flag) {
                if (document.form1.BDID.value == '') {
                    document.form1.BDID.value = BDIDValue;
                }
                else {
                    document.form1.BDID.value += ',' + BDIDValue;
                }
            }
            else {
                if (document.form1.BDID.value.indexOf(',' + BDIDValue) != -1) {
                    document.form1.BDID.value = document.form1.BDID.value.replace(',' + BDIDValue, '')
                }
                else if (document.form1.BDID.value.indexOf(BDIDValue + ',') == 0) {
                    document.form1.BDID.value = document.form1.BDID.value.replace(BDIDValue + ',', '')
                }
                else if (document.form1.BDID.value.indexOf(BDIDValue) == 0) {
                    document.form1.BDID.value = document.form1.BDID.value.replace(BDIDValue, '')
                }
            }
        }
        function ReportPrint() {
            if (document.form1.BDID.value == '') {
                alert('請先勾選要列印的企業');
                return false;
            }
            window.open('../../SQControl.aspx?&SQ_AutoLogout=true&sys=TR&filename=TR_03_001_R1&path=TIMS&BDID=' + document.form1.BDID.value, 'print', 'toolbar=0,location=0,status=0,menubar=0,resizable=1');
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0" class="font">
            <tr>
                <td align="center">
                    <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>首頁&gt;&gt;訓練與就業需求管理&gt;&gt;<font color="#990000">事業單位查詢</font>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="SearchTable" runat="server" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" width="80">公司名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="Uname" runat="server" MaxLength="50" Width="410px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">投保人數
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="KEID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">縣市
                            </td>
                            <td class="whitecol">
                                <asp:CheckBoxList ID="SCTID" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow" RepeatColumns="8">
                                </asp:CheckBoxList>
                                <input id="HidObj" type="hidden" runat="server">
                            </td>
                        </tr>
                    </table>
                    <p align="center">
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button><input id="SearchState" type="hidden" value="1" name="SearchState" runat="server" />
                    </p>
                    <table id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server" class="font">
                        <tr>
                            <td align="right">
                                <asp:LinkButton ID="StateButton" runat="server" ForeColor="Blue">關閉查詢條件</asp:LinkButton>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AutoGenerateColumns="False" Width="100%" AllowPaging="True">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="選取">
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                            <ItemTemplate>
                                                <input id="Checkbox1" type="checkbox" runat="server">
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle Width="25px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Uname" HeaderText="事業單位名稱">
                                            <HeaderStyle Width="150px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Hprsno" HeaderText="投保&lt;BR&gt;人數">
                                            <HeaderStyle Width="30px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Ename" HeaderText="負責人&lt;BR&gt;姓名">
                                            <HeaderStyle Width="50px"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Tel" HeaderText="聯絡電話"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Addr" HeaderText="通訊地址"></asp:BoundColumn>
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
                                <%--
										<asp:Button id="Button2" runat="server" Text="列印清冊"></asp:Button>
                                --%>
                                <asp:Button ID="Button3" runat="server" Text="列印空白企業訪視表" CssClass="asp_Export_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <input id="BDID" type="hidden" runat="server">
    </form>
</body>
</html>
