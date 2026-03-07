<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_002.aspx.vb" Inherits="WDAIIP.CO_01_002" %>

<%--<!DOCTYPE html>--%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>計畫參與度</title>
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
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;計畫參與度</asp:Label>
                </td>
            </tr>
        </table>
        <%--style="display: none"--%>
        <div id="divSch1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <%--<tr>
                    <td class="bluecol_need" width="11%">訓練機構
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="center" runat="server" Width="410px" onfocus="this.blur()"></asp:TextBox>
                        <input id="Org" type="button" value="..." name="Org" runat="server">
                        <input id="RIDValue" style="width: 32px; height: 22px" type="hidden" name="RIDValue" runat="server">
                        <input id="Orgidvalue" style="width: 32px; height: 22px" type="hidden" name="Orgidvalue" runat="server">
                        <span id="HistoryList2" style="position: absolute; display: none">
                            <asp:Table ID="HistoryRID" runat="server" Width="310px">
                            </asp:Table>
                        </span>
                    </td>
                </tr>--%>
                <%-- <tr>
                    <td colspan="4" class="table_title_left">會議場次管理</td>
                </tr>--%>
                <tr>
                    <td class="bluecol" style="width: 20%">轄區
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:CheckBoxList ID="sDistID" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                        </asp:CheckBoxList>
                        <input id="sDistHidden" type="hidden" value="0" name="sDistHidden" runat="server">
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need" style="width: 20%">年度
                    </td>
                    <td class="whitecol" style="width: 30%">
                        <asp:DropDownList ID="SYEARlist" runat="server">
                        </asp:DropDownList>
                    </td>
                    <td class="bluecol_need" style="width: 20%">上／下半年度<%--申請階段--%></td>
                    <td class="whitecol" style="width: 30%">
                        <asp:DropDownList ID="halfYear" runat="server">
                            <asp:ListItem Value="" Selected="True">不區分</asp:ListItem>
                            <asp:ListItem Value="1">上年度</asp:ListItem>
                            <asp:ListItem Value="2">下年度</asp:ListItem>
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" width="11%">訓練機構
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="txtORGNAME" runat="server" Width="410px" MaxLength="100"></asp:TextBox>(查詢關鍵字)
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">匯入總場次</td>
                    <td colspan="3" class="whitecol">
                        <div align="left">
                            <input id="File1" type="file" size="80" name="File1" runat="server" accept=".xls,.ods" />
                            <asp:Button ID="Btn_XlsImport1" runat="server" Text="匯入總場次" CssClass="asp_Export_M"></asp:Button>
                            (必須為ods或xls格式)<asp:HyperLink ID="HyperLink1" runat="server" CssClass="font" ForeColor="#8080FF">下載上載格式檔</asp:HyperLink><%--&nbsp;&nbsp;--%>
                        </div>
                    </td>
                </tr>
                <%-- <tr><td class="whitecol" colspan="4" align="center"></td></tr><tr><td colspan="4" class="table_title_left">會議參與度管理</td></tr>--%>
<%--<tr><td class="bluecol_need" width="11%">匯入訓練機構</td>
<td class="whitecol" colspan="3"><asp:TextBox ID="center" runat="server" Width="410px" onfocus="this.blur()"></asp:TextBox>
<input id="Org" type="button" value="..." name="Org" runat="server" />
<input id="RIDValue" type="hidden" name="RIDValue" runat="server" /></td></tr>--%>
                    <%--<input id="Orgidvalue" type="hidden" name="Orgidvalue" runat="server" />--%>
                    <%--<span id="HistoryList2" style="position: absolute; display: none"><asp:Table ID="HistoryRID" runat="server" Width="310px"></asp:Table></span>--%>
                <tr>
                    <td class="bluecol">匯入各場與會名單</td>
                    <td colspan="3" class="whitecol">
                        <div align="left">
                            <input id="File2" type="file" size="80" name="File2" runat="server" accept=".xls,.ods" />
                            <asp:Button ID="Btn_XlsImport2" runat="server" Text="匯入與會名單" CssClass="asp_Export_M"></asp:Button>
                            (必須為ods或xls格式)
                    <asp:HyperLink ID="HyperLink2" runat="server" CssClass="font" ForeColor="#8080FF">下載上載格式檔</asp:HyperLink><%--&nbsp;&nbsp;--%>
                        </div>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">匯出檔案格式</td>
                    <td colspan="3" class="whitecol">
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
                        <asp:Button ID="btnQuery" runat="server" Text="查詢會議參與度" CssClass="asp_button_S"></asp:Button>
                        <asp:Button ID="BtnSchOP1" runat="server" Text="查詢總場次" CssClass="asp_button_S"></asp:Button>
                        <%--<asp:Button ID="btnImp1" runat="server" Text="匯入總場次" CssClass="asp_button_S"></asp:Button>--%>
                        <asp:Button ID="btnExp1" runat="server" Text="匯出場次代碼" CssClass="asp_Export_M"></asp:Button>
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
                                    <%--<HeaderStyle ForeColor="#00ffff"></HeaderStyle>--%>
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left"
                                        PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                        <Columns>
                                            <asp:BoundColumn HeaderText="序號" HeaderStyle-Width="5%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="Years_ROC" HeaderText="年度" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="halfYearN" HeaderText="上/下<br>半年度" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" />
                                            </asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="OrgName2" HeaderText="管控單位"></asp:BoundColumn>--%>
                                            <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區分署" HeaderStyle-Width="25%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構" HeaderStyle-Width="25%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="CNT1" HeaderText="應出席<br>總場次" HeaderStyle-Width="7%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="CNT2" HeaderText="實際出<br>席場次" HeaderStyle-Width="7%"></asp:BoundColumn>
                                            <%--<asp:BoundColumn DataField="PrjDegree" HeaderText="計畫參與度"></asp:BoundColumn>--%>
                                            <asp:TemplateColumn HeaderText="計畫<br>參與度" HeaderStyle-Width="7%">
                                                <ItemTemplate>
                                                    <asp:Label ID="iLabPrjDegree" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>

                                            <asp:TemplateColumn HeaderText="功能" HeaderStyle-Width="8%">
                                                <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="lbtEdit" runat="server" Text="編輯" CommandName="btnEdit" CssClass="linkbutton"></asp:LinkButton>
                                                    <%--<asp:LinkButton ID="lbtDel1" runat="server" Text="刪除" CommandName="btnDel" CssClass="linkbutton"></asp:LinkButton>--%>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>

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
            </table>
        </div>
        <div id="divEdt1" runat="server">
            <table class="table_sch" cellpadding="1" cellspacing="1" width="100%" border="0">
                <tr>
                    <td class="bluecol" style="width: 20%">訓練機構
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:Label ID="LabOrgName" runat="server"></asp:Label>
                    </td>

                </tr>
                <tr>
                    <td class="bluecol" width="11%">年度
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:Label ID="LabPartyYears" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" width="11%">上/下半年度
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:Label ID="LabhalfYear" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" width="11%">應出席總場次
                    </td>
                    <td class="whitecol" colspan="3">
                        <asp:Label ID="LabShouldTimes" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol" width="11%">實際出席場次
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="LabActTimes" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol" width="11%">計畫參與度
                    </td>
                    <td class="whitecol">
                        <asp:Label ID="LabPrjDegree" runat="server"></asp:Label>
                        <%--&nbsp;<asp:Button ID="btnRecal" runat="server" Text="重計計算" />--%>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="11%">活動場次
                    </td>
                    <td colspan="3" class="whitecol">
                        <asp:CheckBoxList ID="chkbPTYID" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow"></asp:CheckBoxList>
                        <input id="chkbPTYID_hid" type="hidden" value="0" name="chkbPTYID_hid" runat="server" />
                        <%--<HeaderStyle ForeColor="#00ffff"></HeaderStyle>--%>
                    </td>
                </tr>
            </table>
            <table width="100%">
                <tr>
                    <td class="whitecol">
                        <div align="center">
                            <%--<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                            <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="55px">10</asp:TextBox>--%>
                            <asp:Button ID="BtnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                            <asp:Button ID="BtnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
        <asp:HiddenField ID="Hid_YEARS" runat="server" />
        <asp:HiddenField ID="Hid_DISTID" runat="server" />
        <asp:HiddenField ID="Hid_TPLANID" runat="server" />
        <asp:HiddenField ID="Hid_HALFYEAR" runat="server" />
        <asp:HiddenField ID="Hid_ORGID" runat="server" />
    </form>
</body>
</html>
