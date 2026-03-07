<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CO_01_007.aspx.vb" Inherits="WDAIIP.CO_01_007" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>審查計分表開關機制</title>
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
    <%--<script type="text/javascript" language="javascript"></script>--%>
    <style type="text/css">
        .auto-style1 { display: inline-block; padding: 2px 4px; border-radius: 2px; background-color: #0eabd6; color: #FFF; margin-left: 2px; margin-right: 2px; margin-bottom: 2px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;審查計分表&gt;&gt;審查計分表開關機制</asp:Label>
                </td>
            </tr>
        </table>
        <div id="div_sch1" runat="server">
            <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                <tr>
                    <td width="20%" class="bluecol">計畫年度：</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlYEARS_S1" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol">申請階段：</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlAPPSTAGE_S1" runat="server"></asp:DropDownList></td>
                </tr>
                <%--審查計分區間--%>
                <tr>
                    <td class="bluecol">審查計分區間：
                    </td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlTTQSLOCK_S1" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">設定日期：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="schQCDATE1" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span1" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schQCDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span2" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schQCDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="schQCDATE2" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span3" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schQCDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span4" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schQCDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    </td>
                </tr>
                <%--<tr>
                    <td width="20%" class="bluecol">控制起始日期：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="schQSDATE1" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span5" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schQSDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span6" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schQSDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="schQSDATE2" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span7" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schQSDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span8" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schQSDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol">控制結束日期：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="schQFDATE1" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span9" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schQFDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span10" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schQFDATE1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="schQFDATE2" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span11" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= schQFDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span12" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= schQFDATE2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    </td>
                </tr>--%>
                <%--<tr>
                    <td width="20%" class="bluecol">查詢狀態：</td>
                    <td class="whitecol">
                        <asp:CheckBox ID="CHK_ISDELETE" runat="server" Text="包含已刪除" />
                    </td>
                </tr>--%>
            </table>
            <table width="100%">
                <tr>
                    <td class="whitecol" align="center">
                        <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="auto-style1" AuthType="QRY"></asp:Button>
                        <asp:Button ID="bt_add" runat="server" Text="新增" CausesValidation="False" CssClass="asp_button_S" AuthType="ADD"></asp:Button>&nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="center" class="whitecol">
                        <asp:Label ID="lab_msg1" runat="server" ForeColor="Red"></asp:Label></td>
                </tr>
            </table>
            <table id="tb_Sch" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                <tr>
                    <td align="center">
                        <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False">
                            <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                            <HeaderStyle CssClass="head_navy"></HeaderStyle>
                            <Columns>
                                <asp:BoundColumn HeaderText="序號" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="6%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="QCDATE" HeaderText="設定日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="8%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="TTQSLOCK_N" HeaderText="審查計分區間" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="22%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="SPREDATE" HeaderText="(初審)開始日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="11%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="FPREDATE" HeaderText="(初審)結束日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="11%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="QSDATE" HeaderText="單位查詢<br>開放時間" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="11%"></asp:BoundColumn>
                                <asp:BoundColumn DataField="QFDATE" HeaderText="單位查詢<br>結束時間" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="11%"></asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="QEXPLAIN" HeaderText="控制說明" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="30%"></asp:BoundColumn>--%>
                                <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="14%">
                                    <ItemTemplate>
                                        <asp:LinkButton ID="btnEDIT1" runat="server" CommandName="edit" CssClass="asp_button_M">修改</asp:LinkButton>
                                        <%--<asp:LinkButton ID="btnDEL1" runat="server" CommandName="del" CssClass="asp_button_M">刪除</asp:LinkButton>--%>
                                        <%--<asp:Button ID="btnEDIT1" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M" AuthType="UPD" />
                                        <asp:Button ID="btnDEL1" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M" AuthType="DEL" />--%>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                            <PagerStyle Visible="False"></PagerStyle>
                        </asp:DataGrid><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </td>
                </tr>
            </table>
        </div>

        <div id="div_edit" runat="server">
            <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                <tr>
                    <td width="20%" class="bluecol_need">設定日期：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txQCDATE" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span17" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txQCDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span18" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= txQCDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    </td>
                </tr>

                <tr>
                    <td width="20%" class="bluecol_need">計畫年度：</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlYEARS" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol_need">申請階段：</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlAPPSTAGE" runat="server"></asp:DropDownList></td>
                </tr>
                <%--審查計分區間--%>
                <tr>
                    <td class="bluecol_need">審查計分區間：
                    </td>
                    <td class="whitecol">
                        <asp:DropDownList ID="ddlTTQSLOCK" runat="server"></asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" class="table_title">審查計分表(初審)</td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">(初審)開放時間：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txSPREDATE" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span5" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txSPREDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span6" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= txSPREDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                        <asp:DropDownList ID="ddlSPREDATE_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlSPREDATE_MM" runat="server"></asp:DropDownList>分
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">(初審)結束時間：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txFPREDATE" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span7" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txFPREDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span8" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= txFPREDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                        <asp:DropDownList ID="ddlFPREDATE_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlFPREDATE_MM" runat="server"></asp:DropDownList>分
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">作業提醒：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txREMIND1" Width="80%" runat="server" placeholder="此段文字將顯示於首頁【作業提醒】區塊" MaxLength="300" Rows="7" TextMode="MultiLine"></asp:TextBox></td>
                </tr>
                <%--開放單位查詢等級--%>
                <tr>
                    <td colspan="2" class="table_title">開放單位查詢等級</td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">控制起始：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txQSDATE" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span13" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txQSDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span14" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= txQSDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                        <asp:DropDownList ID="ddlQSDATE_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlQSDATE_MM" runat="server"></asp:DropDownList>分
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">控制結束：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txQFDATE" Width="20%" onfocus="this.blur()" runat="server" MaxLength="12"></asp:TextBox>
                        <span id="span15" runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txQFDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                        <span id="span16" runat="server">
                            <img style="cursor: pointer" onclick="javascript:clearDate('<%= txQFDATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                        <asp:DropDownList ID="ddlQFDATE_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddlQFDATE_MM" runat="server"></asp:DropDownList>分
                    </td>
                </tr>
                <tr>
                    <td width="20%" class="bluecol_need">控制說明／作業提醒：</td>
                    <td class="whitecol">
                        <asp:TextBox ID="txQEXPLAIN" Width="80%" runat="server" placeholder="此段文字將顯示於首頁【作業提醒】區塊" MaxLength="300" Rows="7" TextMode="MultiLine"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="whitecol" colspan="2" valign="middle" align="center">
                        <asp:Label ID="LabISDELETE" runat="server" Text="(資料已刪除)" Visible="False" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td class="whitecol" align="center" colspan="2">
                        <asp:Button ID="bt_save" runat="server" Text="儲存" CssClass="asp_button_S" AuthType="SAVE"></asp:Button>
                        <asp:Button ID="bt_cancle" runat="server" Text="取消" CausesValidation="False" CssClass="asp_button_S" AuthType="CANCLE"></asp:Button>&nbsp;
                    </td>
                </tr>
                <tr>
                    <td align="center" class="whitecol" colspan="2">
                        <asp:Label ID="lab_msg2" runat="server" ForeColor="Red"></asp:Label></td>
                </tr>
            </table>

        </div>

        <asp:HiddenField ID="Hid_OTQID" runat="server" />

    </form>
</body>
</html>
