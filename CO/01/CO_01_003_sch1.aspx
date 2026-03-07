<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_003_sch1.aspx.vb" Inherits="WDAIIP.CO_01_003_sch1" %>


<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>TTQS及時勾稽資料</title>
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
            for (var i = 1; i < mytable.rows.length; i++) {
                //var myradio = mytable.rows[i].cells[0].children[0]; //Radio1
                //myradio.checked = false;
                //if (num == i) { myradio.checked = true; }
                var cell = mytable.rows[i].cells[0];
                for (var j = 0; j < cell.childNodes.length; j++) {
                    if (cell.childNodes[j].type == "radio") {
                        cell.childNodes[j].checked = false;
                        if (num == i) {
                            cell.childNodes[j].checked = true;
                        }
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
                                            <td width="20%" class="bluecol">訓練單位：</td>
                                            <td width="60%" class="whitecol">
                                                <asp:Label ID="ORGNAME" runat="server"></asp:Label>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>

                            <tr>
                                <td colspan="2">
                                    <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td width="10%" class="bluecol">
                                                <center>
                                                    TTQS<br>
                                                    評核結果</center>
                                            </td>
                                            <td align="center">
                                                <%--評核版別	申請目的	評核結果	展延	評核日期	發文日期	有效期限	備註--%>
                                                <%--[SENDVER]	[GOAL]	[RESULT]	[EXTLICENS]	[SENDDATE]	[ISSUEDATE]	[VALIDDATE]	[MEMO]--%>
                                                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" Width="100%" CssClass="font" CellPadding="8">
                                                    <AlternatingItemStyle BackColor="#f5f5f5" />
                                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:TemplateColumn>
                                                            <HeaderStyle Width="5%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <input id="Radio1" type="radio" value="Radio1" runat="server" />
                                                                <asp:HiddenField ID="Hid_VTSID" runat="server" />
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:BoundColumn DataField="ROWNUM" HeaderText="序號">
                                                            <HeaderStyle Width="6%"></HeaderStyle>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="SENDVER_N" HeaderText="評核版別">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="GOAL" HeaderText="申請目的">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="RESULT_N" HeaderText="評核結果">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="EXTLICENS" HeaderText="展延">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="SENDDATE" HeaderText="評核日期">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="ISSUEDATE" HeaderText="發文日期">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="VALIDDATE" HeaderText="有效期限">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <asp:BoundColumn DataField="MEMO2" HeaderText="TTQS訓練機構名稱">
                                                            <ItemStyle HorizontalAlign="Center" />
                                                        </asp:BoundColumn>
                                                        <%--<asp:TemplateColumn>
                                                            <HeaderTemplate>備註</HeaderTemplate>
                                                            <ItemStyle HorizontalAlign="Center" />
                                                            <ItemTemplate>
                                                                <asp:Label ID="lbMEMO" runat="server"></asp:Label></ItemTemplate>
                                                        </asp:TemplateColumn>--%>
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
            <%--function open_CO01003sch1(s_OTTID, s_ORGID, s_COMIDNO) --%>
            <asp:HiddenField ID="Hid_comidno" runat="server" />
            <asp:HiddenField ID="Hid_ORGID" runat="server" />
            <asp:HiddenField ID="Hid_OTTID" runat="server" />
            <asp:HiddenField ID="Hid_RESULT1" runat="server" />
        </form>
    </div>
</body>
</html>
