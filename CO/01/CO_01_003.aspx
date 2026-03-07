<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_003.aspx.vb" Inherits="WDAIIP.CO_01_003" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>最近一次TTQS評核結果等級</title>
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
        //CheckboxAll
        function ChangeAll(obj) {
            var objLen = document.form1.length;
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (document.form1.elements[iCount].type == "checkbox") {
                    var mycheck = document.form1.elements[iCount];
                    if (!mycheck.disabled) {
                        mycheck.checked = (obj.checked == true ? true : false);
                    }
                }
            }
        }

        //更新專用
        function open_CO01003sch1(s_OTTID, s_ORGID, s_COMIDNO) {
            var msg = "";
            var rqID = getParamValue('ID');
            if (s_OTTID == '') { msg += '請輸入流水TTQS號碼!\n'; }
            if (s_ORGID == '') { msg += '請輸入流水機構號碼!\n'; }
            if (s_COMIDNO == '') { msg += '請輸入機構統編號碼!\n'; }
            if (msg != "") {
                alert(msg);
                return false; //結束。
            }
            //DataGridTable
            document.getElementById("DataGridTable").style.display = "none";
            var url1 = "";
            url1 = "CO_01_003_sch1.aspx?ID=" + rqID
            + "&SPAGE=CO01003"
            + "&OTTID=" + s_OTTID
            + "&ORGID=" + s_ORGID
            + "&COMIDNO=" + s_COMIDNO
            //+ "&TK=" + escape(s_TK)
            wopen(url1, 'WNCO01003sch', 1200, 680, 1);

            return false;
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;最近一次TTQS評核結果等級</asp:Label>

                </td>
            </tr>
        </table>
        <%--style="display: none"--%>
        <div id="divSch1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol" style="width: 20%">分署
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="ddlDISTID" runat="server"></asp:DropDownList>
                    </td>
                </tr>

                <tr>
                    <td class="bluecol_need" style="width: 20%">年度</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="SYEARlist" runat="server"></asp:DropDownList></td>
                    <td class="bluecol_need" style="width: 20%">截止時間</td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rblMONTHS" runat="server"></asp:RadioButtonList></td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 20%">訓練機構
                    </td>
                    <td class="whitecol" style="width: 30%">
                        <asp:TextBox ID="OrgName" runat="server" MaxLength="50" Columns="60" Width="80%"></asp:TextBox>
                        <%--
                        <asp:TextBox ID="center" runat="server" Width="410px" onfocus="this.blur()"></asp:TextBox>
                        <input id="Org" type="button" value="..." name="Org" runat="server">
                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                            <input id="Orgidvalue" type="hidden" name="Orgidvalue" runat="server">
                            <span id="HistoryList2" style="position: absolute; display: none">
                            <asp:Table ID="HistoryRID" runat="server" Width="310px">
                            </asp:Table>
                        </span>--%>
                    </td>
                    <td class="bluecol" style="width: 20%">統一編號
                    </td>
                    <td class="whitecol" style="width: 30%">
                        <asp:TextBox ID="COMIDNO" runat="server" MaxLength="15" Width="50%"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                </tr>
                <tr id="TRPlanPoint28" runat="server">
                    <td class="bluecol_need">計畫</td>
                    <td class="whitecol" colspan="3">
                        <asp:RadioButtonList ID="OrgPlanKind" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                            <asp:ListItem Value="G" Selected="True">產業人才投資計畫</asp:ListItem>
                            <asp:ListItem Value="W">提升勞工自主學習計畫</asp:ListItem>
                        </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">機構別 </td>
                    <td class="whitecol">
                        <asp:DropDownList ID="OrgKindList" runat="server" CssClass="font">
                        </asp:DropDownList>
                    </td>
                    <td class="bluecol_need">單位確認 </td>
                    <td class="whitecol">
                        <asp:RadioButtonList ID="rblCONFIRM" runat="server"></asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">評核版本</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlSENDVER" runat="server" CssClass="font">
                        </asp:DropDownList>
                    </td>
                    <td class="bluecol">評核結果</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlRESULT" runat="server" CssClass="font">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td class="whitecol" colspan="3">
                    <asp:RadioButtonList ID="RBListExpType" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                        <asp:ListItem Value="EXCEL" Selected="True">EXCEL</asp:ListItem>
                        <asp:ListItem Value="ODS">ODS</asp:ListItem>
                    </asp:RadioButtonList>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" colspan="4" align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>
                        <%--<asp:Button ID="btnImp1" runat="server" Text="匯入總場次" CssClass="asp_button_S"></asp:Button>--%>
                        <asp:Button ID="btnExp1" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" colspan="4" align="center">
                        <asp:Label ID="labmmo1" runat="server" ForeColor="Red">※【訓練單位確認結果】滑鼠移到上方，可顯示資料有誤的原因!</asp:Label>
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
                                            <asp:TemplateColumn HeaderText="選取">
                                                <HeaderStyle Width="5%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center" />
                                                <HeaderTemplate>選取<input id="CheckboxAll" type="checkbox" runat="server" /></HeaderTemplate>
                                                <ItemTemplate>
                                                    <input id="checkbox1" type="checkbox" runat="server">
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn HeaderText="序號">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="YEARS_ROC" HeaderText="年度"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="DISTNAME" HeaderText="分署"></asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="HALFYEARN" HeaderText="上/下<br>半年度"></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="機構名稱"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="COMIDNO" HeaderText="統一編號"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ORGKIND_N" HeaderText="機構別"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="SENDVER_N" HeaderText="評核版別"></asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="GOAL" HeaderText="申請目的"></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="RESULT_N" HeaderText="評核結果"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="EXTLICENS" HeaderText="展延"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="SENDDATE" HeaderText="評核日期"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ISSUEDATE" HeaderText="發文日期"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="VALIDDATE" HeaderText="有效期限"></asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="MEMO" HeaderText="備註"></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="APPLIEDRESULT_N" HeaderText="審核狀況"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="MEMO2" HeaderText="TTQS訓練機構名稱"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="CONFIRM_N" HeaderText="訓練單位確認結果"></asp:BoundColumn>
                                            <%-- <asp:TemplateColumn HeaderText="功能">
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lbtEdit" runat="server" Text="編輯" CommandName="btnEdit" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>--%>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <ItemTemplate>
                                                    <asp:HiddenField ID="Hid_OTTID" runat="server" />
                                                    <asp:LinkButton ID="lbtRENEW" runat="server" Text="更新" CommandName="RENEW" CssClass="linkbutton"></asp:LinkButton>
                                                    <br />
                                                    <asp:LinkButton ID="lbtUNLOCK" runat="server" Text="解鎖" CommandName="UNLOCK" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>

                                        </Columns>
                                        <PagerStyle Visible="False" HorizontalAlign="Left" ForeColor="Blue" Position="Top" Mode="NumericPages"></PagerStyle>
                                    </asp:DataGrid>
                                </td>
                            </tr>
                            <tr>
                                <td class="whitecol" align="center">
                                    <asp:Button ID="BtnSaveData1" runat="server" Text="審核確認" CssClass="asp_button_S"></asp:Button>
                                    <%--<div align="center"></div>--%>
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


            </table>
        </div>

        <%--<asp:HiddenField ID="Hid_OTSID" runat="server" />
        <asp:HiddenField ID="Hid_RTSID" runat="server" />--%>
    </form>
</body>
</html>
