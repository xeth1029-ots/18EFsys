<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_001_sch2.aspx.vb" Inherits="WDAIIP.SD_01_001_sch2" %>


<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>勞保勾稽資料</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function checkRadio(num) {
            var mytable = document.getElementById('DataGrid1');
            var cell;
            for (var i = 1; i < mytable.rows.length; i++) {
                //var myradio = mytable.rows[i].cells[0].children[0]; //Radio1
                //myradio.checked = false;
                //if (num == i) { myradio.checked = true; }
                cell = mytable.rows[i].cells[0];
                for (j = 0; j < cell.childNodes.length; j++) {
                    if (cell.childNodes[j].type == "radio") {
                        cell.childNodes[j].checked = false;
                        if (num == i) { cell.childNodes[j].checked = true; }
                    }
                }
            }
        }
    </script>
</head>
<body>
    <div style="overflow-y: auto; height: 650px;">
        <form id="form1" method="post" runat="server">
            <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td colspan="2">
                                    <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td width="20%" class="bluecol">姓名：<asp:Label ID="labNAME" runat="server"></asp:Label></td>
                                            <td width="20%" class="bluecol">身分證號：<asp:Label ID="labIDNO" runat="server"></asp:Label></td>
                                            <td width="60%" class="whitecol">
                                                <asp:Button ID="btnPrint" runat="server" Text="列印" OnClientClick="window.print();return false;" CssClass="asp_Export_M"></asp:Button>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <%--<tr>
                                <td colspan="2">
                                    <table class="table_sch" id="Table3">
                                        <tr>
                                            <td width="10%" rowspan="6" class="bluecol">
                                                <center>就服中心</center>
                                            </td>
                                            <td class="whitecol" style="background-color: #ecf7ff">輔導就業狀況：<asp:Label ID="WorkState" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" style="background-color: #ecf7ff">廠商名稱：<asp:Label ID="COMPNAME" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" style="background-color: #ecf7ff">地址：<asp:Label ID="COMPADDR" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" style="background-color: #ecf7ff">電話：<asp:Label ID="CONTEL" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" style="background-color: #ecf7ff">錄用月薪：<asp:Label ID="REPLY_SALARY" runat="server"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td class="whitecol" style="background-color: #ecf7ff">就業日期：<asp:Label ID="REPLY_WKDATE" runat="server"></asp:Label></td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>--%>
                            <tr>
                                <td colspan="2">
                                    <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td width="10%" class="bluecol">
                                                <center>勞保局資料</center>
                                            </td>
                                            <td align="center">
                                                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="font" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:TemplateColumn>
                                                            <HeaderStyle Width="5%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <input id="Radio1" type="radio" value="Radio1" runat="server" />
                                                                <asp:HiddenField ID="Hid_sb4id" runat="server" />
                                                                <asp:HiddenField ID="Hid_SMDATE" runat="server" />
                                                                <asp:HiddenField ID="Hid_FMDATE" runat="server" />
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:BoundColumn HeaderText="序號">
                                                            <HeaderStyle Width="6%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:TemplateColumn>
                                                            <HeaderTemplate>投保種類</HeaderTemplate>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Label ID="LabActNoType" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <%--<asp:BoundColumn HeaderText="投保種類"></asp:BoundColumn>--%>
                                                        <asp:BoundColumn DataField="ActNo" HeaderText="保險證號">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="MDate" HeaderText="異動日期">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <%--<asp:BoundColumn DataField="ChangeMode" HeaderText="異動狀況"></asp:BoundColumn>--%>
                                                        <asp:TemplateColumn>
                                                            <HeaderTemplate>異動狀況</HeaderTemplate>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Label ID="LabChangeMode" runat="server"></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <%--<asp:BoundColumn DataField="Salary" HeaderText="投保薪資級距">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>--%>
                                                        <asp:BoundColumn DataField="UName" HeaderText="投保單位"></asp:BoundColumn>
                                                        <%--<asp:BoundColumn DataField="deptmentN" HeaderText="公法救助"></asp:BoundColumn>--%>
                                                        <asp:BoundColumn DataField="biefN" HeaderText="公法救助"></asp:BoundColumn>
                                                        <asp:TemplateColumn>
                                                            <HeaderTemplate>ECFA身分</HeaderTemplate>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Label ID="LabECFA" runat="server"></asp:Label></ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:BoundColumn DataField="BIEFDESC" HeaderText="註記"></asp:BoundColumn>
                                                    </Columns>
                                                </asp:DataGrid>
                                                <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2" class="whitecol">
                                    <asp:Button ID="Button1" runat="server" Text="確定" CssClass="asp_button_M"></asp:Button>
                                    <input id="Button2" type="button" value="離開" name="Button2" runat="server" class="asp_button_M">
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            <asp:HiddenField ID="Hid_idno" runat="server" />
            <asp:HiddenField ID="Hid_birth" runat="server" />
            <asp:HiddenField ID="Hid_CNAME" runat="server" />
            <asp:HiddenField ID="Hid_SPAGE" runat="server" />
            <asp:HiddenField ID="Hid_ECFA_YES" runat="server" />
        </form>
    </div>
</body>
</html>
