<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_02_002.aspx.vb" Inherits="WDAIIP.TC_02_002" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>申復申請作業</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        //[全選／全不選] //chkItem
        function doSelectAll(obj) {
            $("input[type=checkbox][data-role='chkItem']:enabled").prop("checked", obj.checked);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" cellspacing="1" id="FrameTable" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="font">
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;申復申請作業</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="panelSch" runat="server">
            <table class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" width="16%">訓練機構</td>
                                <td colspan="3" class="whitecol" width="84%">
                                    <asp:TextBox ID="center" runat="server" Width="60%" ></asp:TextBox><%--onfocus="this.blur()"--%>
                                    <input id="Org" type="button" value="..." name="Org" runat="server">
                                    <input id="RIDValue" style="width: 10%;" type="hidden" name="RIDValue" runat="server">
                                    <input id="Orgidvalue" style="width: 10%;" type="hidden" name="Orgidvalue" runat="server">
                                    <span id="HistoryList2" style="position: absolute; display: none">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級名稱</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="ClassName" runat="server" Columns="44" MaxLength="55" Width="88%"></asp:TextBox></td>
                                <td class="bluecol" width="16%">期別</td>
                                <td class="whitecol" width="34%">
                                    <asp:TextBox ID="CyclType" runat="server" Columns="10" MaxLength="3" Width="33%"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">申請階段</td>
                                <td class="whitecol" colspan="3">
                                    <asp:DropDownList ID="ddlAPPSTAGE_SCH" runat="server"></asp:DropDownList></td>
                            </tr>
                            <%--增加【轉班上架】欄位，選項：不區分、未轉班、已轉班--%>
                            <tr id="tr_TransFlag" runat="server">
                                <td class="bluecol">轉班上架</td>
                                <td colspan="3" class="whitecol">
                                    <asp:RadioButtonList ID="rbl_TransFlagS" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                                        <asp:ListItem Value="A" Selected="True">不區分</asp:ListItem>
                                        <asp:ListItem Value="N">未轉班</asp:ListItem>
                                        <asp:ListItem Value="Y">已轉班</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                        <table cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="whitecol">
                                    <div align="center">
                                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                        <asp:Button ID="btnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                        <%--<asp:Button ID="btnONLINE1" runat="server" Text="線上申辦" CssClass="asp_Export_M"></asp:Button>--%>
                                    </div>
                                </td>
                            </tr>
                        </table>
                        <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                            <tr>
                                <td align="center">
                                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                        <tr>
                                            <td>訓練計畫：<asp:Label ID="TPlanName" runat="server" CssClass="font"></asp:Label></td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:DataGrid ID="dtPlan" runat="server" Width="100%" CssClass="font" AllowSorting="True" PagerStyle-HorizontalAlign="Left" PagerStyle-Mode="NumericPages" AllowPaging="True" AutoGenerateColumns="False">
                                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                                    <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                                    <Columns>
                                                        <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="YEARSROCAG" HeaderText="計畫年度" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="APPLIEDDATE" HeaderText="申請日期" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="STDATE" HeaderText="訓練起日" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="FTDATE" HeaderText="訓練迄日" ItemStyle-HorizontalAlign="Center"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="DISTNAME" HeaderText="管控單位"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="機構名稱"></asp:BoundColumn>
                                                        <asp:BoundColumn DataField="CLASSCNAME" HeaderText="班名"></asp:BoundColumn>
                                                        <asp:TemplateColumn HeaderText="審核狀態">
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:Label ID="labAppliedResult" runat="server" Text=""></asp:Label>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="功能">
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:LinkButton ID="lbtSFEDIT1" runat="server" Text="申復" CommandName="lbtSFEDIT1" CssClass="linkbutton"></asp:LinkButton>
                                                                <asp:LinkButton ID="lbtSFDEL1" runat="server" Text="刪除" CommandName="lbtSFDEL1" CssClass="linkbutton"></asp:LinkButton>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                        <asp:TemplateColumn HeaderText="列印">
                                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                            <ItemTemplate>
                                                                <asp:LinkButton ID="lbtSFPRINT1" runat="server" Text="申復意見表" CommandName="lbtSFPRINT1" CssClass="linkbutton"></asp:LinkButton>
                                                            </ItemTemplate>
                                                        </asp:TemplateColumn>
                                                    </Columns>
                                                    <PagerStyle Visible="False" HorizontalAlign="Left" ForeColor="Blue" Position="Top" Mode="NumericPages"></PagerStyle>
                                                </asp:DataGrid>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="center">
                                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td align="center">
                                    <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="panelEdit1" runat="server">
            <table id="tbPanelEdit1" class="table_sch" border="0" cellspacing="1" cellpadding="1" width="100%">
                <tr>
                    <td class="bluecol" width="16%">年度 </td>
                    <td class="whitecol" width="34%">
                        <asp:Label ID="lbYEARS_ROC" runat="server"></asp:Label>
                        <asp:Label ID="lbAPPSTAGE_N" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol" width="16%">轄區 </td>
                    <td class="whitecol" width="34%">
                        <asp:Label ID="lbDistName" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">訓練機構 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbOrgName" runat="server"></asp:Label></td>
                    <td class="bluecol">課程申請流水號 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbPSNO28" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">班級名稱 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbClassName" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol">訓練期間 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbSFTDate" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">職類課程</td>
                    <td class="whitecol">
                        <asp:Label ID="lbGCODEPNAME" runat="server"></asp:Label>
                        <%--訓練業別<asp:Label ID="lbGCNAME" runat="server"></asp:Label>--%>
                    </td>
                    <td class="bluecol">訓練職能 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbCCNAME" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">訓練人次 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbTNum" runat="server"></asp:Label>
                    </td>
                    <td class="bluecol">訓練時數 </td>
                    <td class="whitecol">
                        <asp:Label ID="lbTHours" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="table_title" colspan="4">申復資料</td>
                </tr>
                <tr>
                    <td class="bluecol_need">聯絡人 </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SFCONTNAME" runat="server"></asp:TextBox><%--<asp:Label ID="lbSFCONTNAME" runat="server" Text=""></asp:Label>--%>
                    </td>
                    <td class="bluecol_need">聯絡電話 </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SFCONTTEL" runat="server"></asp:TextBox><%--<asp:Label ID="lbSFCONTTEL" runat="server" Text=""></asp:Label>--%>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">職稱 </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SFCONTTITLE" runat="server"></asp:TextBox>
                    </td>
                    <td class="bluecol_need">EMAIL </td>
                    <td class="whitecol">
                        <asp:TextBox ID="SFCONTEMAIL" runat="server"></asp:TextBox><%--<asp:Label ID="lbSFCONTEMAIL" runat="server" Text=""></asp:Label>--%>
                    </td>
                </tr>
                <tr>
                    <%--<td class="TC_TD3" align="right">未核班原因</td>--%>
                    <td class="TC_TD3" align="right">審查意見</td>
                    <td class="whitecol" colspan="3">
                        <%--<asp:Label ID="labNGREASON" runat="server" Text=""></asp:Label>--%>
                        <%--<asp:TextBox ID="NGREASON" runat="server" TextMode="multiline" Width="88%" Rows="20">(無)</asp:TextBox>--%>
                        <asp:TextBox ID="NGREASON" runat="server" Rows="20" TextMode="multiline" Width="88%"></asp:TextBox>
                        <%--<br />--%>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol_need">申復理由及說明</td>
                    <td class="whitecol" colspan="3">
                        <asp:TextBox ID="SFCONTREASONS" runat="server" TextMode="multiline" Width="88%" Rows="40"></asp:TextBox></td>
                </tr>
                <tr>
                    <td colspan="4" align="center" class="whitecol">
                        <asp:Button ID="btnSAVE1" runat="server" Text="儲存確認" CssClass="asp_button_M"></asp:Button>&nbsp;
						<asp:Button ID="btnBACK1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <input id="isBlack" type="hidden" name="isBlack" runat="server" />
        <input id="orgname" type="hidden" name="orgname" runat="server" />
        <asp:HiddenField ID="Hid_PSOID" runat="server" />
        <asp:HiddenField ID="Hid_PSNO28" runat="server" />
        <%-- <asp:HiddenField ID="hid_PPINFOtable_guid1" runat="server" /> <asp:HiddenField ID="hid_TPlanID" runat="server" />--%>
    </form>
</body>
</html>
