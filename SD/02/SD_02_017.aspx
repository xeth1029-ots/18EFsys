<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_017.aspx.vb" Inherits="WDAIIP.SD_02_017" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>重大災害提醒處理</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button6').click();
        }
        function SetOneOCID() {
            document.getElementById('Button7').click();
        }
        function choose_class() {
            //var RID = document.form1.RIDValue.value;
            var RIDValue = document.getElementById('RIDValue').value;
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue);
        }
        function search() {
            var OCIDValue1 = document.getElementById('OCIDValue1').value;
            if (OCIDValue1 == '') {
                alert('請選擇職類班別!');
                return false;
            }
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <asp:Button ID="Button7" Style="display: none" runat="server"></asp:Button>
        <asp:Button ID="Button6" Style="display: none" runat="server"></asp:Button>
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
								首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;<font color="#990000">重大災害提醒處理</font>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                        <tr>
                            <td width="100" class="bluecol">訓練機構 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" runat="server">
                                <input id="Button5" onclick="javascript: wopen('../../Common/LevOrg1.aspx', '訓練機構', 300, 300, 1)" type="button" value="..." runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                    </asp:Table>
                                </span></td>
                        </tr>
                        <tr>
                            <td width="100" class="bluecol">職類/班別 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" runat="server">
                                <input id="OCIDValue1" type="hidden" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="310">
                                    </asp:Table>
                                </span></td>
                        </tr>
                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol">
                                <p align="center">
                                    <asp:Button ID="BtnSearch1" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;&nbsp;
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="center">
                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                    </table>
                    <br />
                    <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="40px" VerticalAlign="middle"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="報名機構">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="報名班級">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="已轉知">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                已轉知
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:HiddenField ID="hid_OCID1" runat="server" />
                                                <asp:HiddenField ID="hid_IDNO" runat="server" />
                                                <asp:HiddenField ID="hid_SETID" runat="server" />
                                                <asp:HiddenField ID="hid_eSETID" runat="server" />
                                                <asp:CheckBox ID="CheckBox1" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <p align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:Button ID="btnSave1" runat="server" Text="儲存" CssClass="asp_button_S" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
