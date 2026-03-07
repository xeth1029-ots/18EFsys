<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_10_005_R.aspx.vb" Inherits="WDAIIP.SD_10_005_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>外籍人士上課時數證明</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function search1() {
            var msg = '';
            if (document.form1.OCIDValue1.value == '') msg += '請選擇職類班別\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }

        function choose_class() {
            if (document.form1.RIDValue) {
                var RID = document.form1.RIDValue.value;
                openClass('../02/SD_02_ch.aspx?RID=' + RID);
            }
            else {
                openClass('../02/SD_02_ch.aspx');
            }
        }
        function GETvalue() { document.getElementById('Button7').click(); }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;證書及證明管理&gt;&gt;外籍人士上課時數證明</asp:Label>
                </td>
            </tr>
        </table>
        <table id="Frame" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
            <tr>
                <td align="center">
                    <table id="Table3" class="table_sch" cellpadding="1" cellspacing="1" width="100%">
                        <tr id="tr_center" runat="server">
                            <td class="bluecol" style="width: 20%">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button5" type="button" value="..." runat="server" class="button_b_Mini" />
                                <asp:Button ID="Button7" Style="display: none" runat="server" Text="Button7"></asp:Button>
                                <span id="HistoryList2" style="z-index: 100; position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" style="width: 20%">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" Width="30%" onfocus="this.blur()"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" Width="30%" onfocus="this.blur()"></asp:TextBox>
                                <input onclick="javascript: openClass('../02/SD_02_ch.aspx')" type="button" value="..." name="Submit">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="display: none; left: 260px; position: absolute">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                    </table>
                    <p align="center" class="whitecol">
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    </p>
                    <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server" class="font">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DG_stud" runat="server" Width="100%" AutoGenerateColumns="False" AllowPaging="True" AllowSorting="True" AllowCustomPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy" Width="20%"></HeaderStyle>
                                    <ItemStyle HorizontalAlign="Center" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="選取">
                                            <ItemTemplate>
                                                <input id="Checkbox1" type="checkbox" runat="server" name="Checkbox1">
                                                <asp:HiddenField ID="Hid_StudentID" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="StudentID" HeaderText="學號"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="EngName" HeaderText="英文姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="StudStatus" HeaderText="學員狀態"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="StudentID" HeaderText="StudentID"></asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="submit" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
