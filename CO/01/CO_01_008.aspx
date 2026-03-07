<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_008.aspx.vb" Inherits="WDAIIP.CO_01_008" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>審查計分表(核定)</title>
    <%--<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />--%>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-confirm.min.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery.blockUI.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/TIMS.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        $(document).ready(function () {
            // 初始化 DropDownList 的值 ddlSECONDCHK_ALL
            var ddlSECONDCHK_ALL = $(".csdatagrid1_ddlsecondchk_all");
            ddlSECONDCHK_ALL.val("");
            // 綁定事件
            ddlSECONDCHK_ALL.change(function () {
                // 獲取所有 ddlFIRSTCHK 元素
                var ddlSECONDCHKs = $("#DataGrid1").find(".cssecondchk:not(:disabled)");
                if (ddlSECONDCHKs.length == 0) { return; }
                // 遍歷所有元素
                for (var i = 0; i < ddlSECONDCHKs.length; i++) {
                    ddlSECONDCHKs[i].value = $(this).val();
                }
            });
        });

        function chackAll() {
            var Mytable = document.getElementById('DataGrid1');
            var jChoose1 = $('#Choose1');
            for (var i = 1; i < Mytable.rows.length; i++) {
                var mycheck = Mytable.rows[i].cells[0].children[0];
                if (!mycheck.disabled) {
                    mycheck.checked = jChoose1.prop("checked");//document.form1.Choose1.checked;
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;審查計分表(核定)</asp:Label>
                </td>
            </tr>
        </table>
        <%--style="display: none"--%>
        <div id="divSch1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol_need">分署
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlDISTID" runat="server"></asp:DropDownList>
                        <%--<asp:Label ID="lab_IMPDIST_MSG" runat="server" Text="(匯入必選)"></asp:Label>--%>
                    </td>
                </tr>
                <%--審查計分區間--%>
                <tr>
                    <td class="bluecol_need">審查計分區間
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlSCORING" runat="server"></asp:DropDownList>
                        <%--<asp:Label ID="lab_IMPSCORING_MSG" runat="server" Text="(匯入必選)"></asp:Label>--%>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 16%">訓練機構</td>
                    <td class="whitecol" style="width: 48%">
                        <asp:TextBox ID="OrgName" runat="server" MaxLength="50" Columns="60" Width="80%"></asp:TextBox>
                        <%--<asp:TextBox ID="center" runat="server" Width="410px" onfocus="this.blur()"></asp:TextBox><input id="Org" type="button" value="..." name="Org" runat="server">
    <input id="RIDValue" type="hidden" name="RIDValue" runat="server"><input id="Orgidvalue" type="hidden" name="Orgidvalue" runat="server">
    <span id="HistoryList2" style="position: absolute; display: none"><asp:Table ID="HistoryRID" runat="server" Width="310px"></asp:Table></span>--%>
                    </td>
                    <td class="bluecol" style="width: 16%">統一編號
                    </td>
                    <td class="whitecol" style="width: 20%">
                        <asp:TextBox ID="COMIDNO" runat="server" MaxLength="15" Width="50%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" colspan="4" align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <%--<asp:Button ID="btnImp1" runat="server" Text="匯入總場次" CssClass="asp_button_S"></asp:Button>--%>
                        <%--<asp:Button ID="btnExp1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>--%>
                    </td>
                </tr>
            </table>
            <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center">
                        <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td align="center">
                                    <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left"
                                        PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號" HeaderStyle-Width="7%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="DISTNAME3" SortExpression="DISTID" HeaderText="轄區分署" HeaderStyle-Width="8%">
                                                <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="訓練單位" HeaderStyle-Width="16%" SortExpression="ORGNAME">
                                                <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labORGNAME" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <%--<asp:BoundColumn DataField="ORGNAME" HeaderText="訓練單位" HeaderStyle-Width="16%"></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="COMIDNO" HeaderText="統一編號" HeaderStyle-Width="7%" SortExpression="COMIDNO">
                                                <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="等級" HeaderStyle-Width="7%" SortExpression="RLEVEL_2">
                                                <HeaderStyle ForeColor="#00ffff"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" CssClass="whitecol" />
                                                <ItemTemplate>
                                                    <asp:Label ID="lRLEVEL_2X" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <%--<asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="7%"><ItemStyle HorizontalAlign="Center" Font-Size="Small" /><ItemTemplate>
    <asp:LinkButton ID="lbtView" runat="server" Text="檢視" CommandName="btnView" CssClass="linkbutton"></asp:LinkButton></ItemTemplate></asp:TemplateColumn>--%>                                            <%--<asp:BoundColumn DataField="MEMO" HeaderText="備註"></asp:BoundColumn>--%>
                                        </Columns>
                                        <PagerStyle Visible="False" HorizontalAlign="Left" ForeColor="Blue" Position="Top" Mode="NumericPages"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>

                        </table>
                    </td>
                </tr>
                <tr>
                    <td align="center">
                        <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol">
                        <div align="center">
                            <%--<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="55px">10</asp:TextBox>
                            <asp:Button ID="BtnBack1" runat="server" Text="回上頁" CssClass="asp_button_S"></asp:Button>--%>
                            <%--<asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>--%>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
