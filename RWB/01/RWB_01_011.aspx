<%@ Page AspCompat="true" Language="vb" AutoEventWireup="true" CodeBehind="RWB_01_011.aspx.vb" Inherits="WDAIIP.RWB_01_011" EnableEventValidation="false" %>

<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>首頁資料BANNER</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css" />
    <%--<script type="text/javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //檢查檔案上傳位置
        function CheckAddPIC() {
            var msg = '';
            if (document.getElementById('File1').value == '') msg += '請輸入檔案上傳位置\n';
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;報名網維護&gt;&gt;首頁資料BANNER</asp:Label>
                </td>
            </tr>
        </table>
        <br>
        <table id="tb_SchV" runat="server" class="table_nw" width="100%" cellpadding="1" cellspacing="1">
            <tr>
                <td width="20%" class="bluecol">起始日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="START_DATE_S1" Width="20%" onfocus="this.blur()" runat="server" placeholder="請輸入 起始日期查詢" MaxLength="11"></asp:TextBox>
                    <span id="span1" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= START_DATE_S1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span2" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= START_DATE_S1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="START_DATE_S2" Width="20%" onfocus="this.blur()" runat="server" MaxLength="11"></asp:TextBox>
                    <span id="span3" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= START_DATE_S2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span4" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= START_DATE_S2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol">結束日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="END_DATE_S1" Width="20%" onfocus="this.blur()" runat="server" placeholder="請輸入 結束日期查詢" MaxLength="11"></asp:TextBox>
                    <span id="span5" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= END_DATE_S1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span6" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= END_DATE_S1.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span> ～
                    <asp:TextBox ID="END_DATE_S2" Width="20%" onfocus="this.blur()" runat="server" MaxLength="11"></asp:TextBox>
                    <span id="span7" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= END_DATE_S2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span8" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= END_DATE_S2.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">啟用狀態：</td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="rblISUSE_S" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                        <asp:ListItem Value="X" Selected="True">全部不區分</asp:ListItem>
                        <asp:ListItem Value="Y">啟用</asp:ListItem>
                        <asp:ListItem Value="N">(停用)</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="whitecol" align="center" colspan="2">
                    <asp:Button ID="btnSearch1" runat="server" Text="查詢" CssClass="asp_button_S" AuthType="QRY"></asp:Button>
                    &nbsp;<asp:Button ID="btnAdd1" runat="server" Text="新增" CssClass="asp_button_S" AuthType="ADD"></asp:Button>&nbsp;
                </td>
            </tr>
            <tr>
                <td class="whitecol" align="center" colspan="2">
                    <asp:Label ID="msg1" runat="server" ForeColor="Red"></asp:Label></td>
            </tr>
        </table>

        <table id="tb_Sch" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" Width="100%" AutoGenerateColumns="False">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn DataField="BANNERID" HeaderText="序號" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="6%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="START_DATE" HeaderText="起始日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="END_DATE" HeaderText="結束日期" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="B_TITLE" HeaderText="抬頭" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                            <%--<asp:BoundColumn DataField="B_CONTENT" HeaderText="內容" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>--%>
                            <asp:BoundColumn DataField="B_ALT" HeaderText="提示訊息" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="FILE_NAME" HeaderText="檔名" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="ISUSED_N" HeaderText="啟用" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="6%"></asp:BoundColumn>
                            <asp:BoundColumn DataField="SEQ" HeaderText="順序" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="6%"></asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="14%">
                                <ItemTemplate>
                                    <asp:Button ID="btnUPD1" runat="server" Text="修改" CommandName="UPD1" CssClass="asp_button_M" AuthType="UPD" />
                                    <%--<asp:Button ID="btnDEL1" runat="server" Text="刪除" CommandName="DEL1" CssClass="asp_button_M" AuthType="DEL" />--%>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>

        <table id="tb_Edit1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td width="20%" class="bluecol_need">起始日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="START_DATE" Width="20%" onfocus="this.blur()" runat="server" placeholder="請輸入 起始日期" MaxLength="11"></asp:TextBox>
                    <span id="span9" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= START_DATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span10" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= START_DATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    <asp:DropDownList ID="ddl_SDATE_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddl_SDATE_MM" runat="server"></asp:DropDownList>分
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">結束日期：</td>
                <td class="whitecol">
                    <asp:TextBox ID="END_DATE" Width="20%" onfocus="this.blur()" runat="server" placeholder="請輸入 結束日期" MaxLength="11"></asp:TextBox>
                    <span id="span11" runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= END_DATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    <span id="span12" runat="server">
                        <img style="cursor: pointer" onclick="javascript:clearDate('<%= END_DATE.ClientID %>');" alt="" align="top" src="../../images/CloseMsn.gif" width="20" height="20" /></span>
                    <asp:DropDownList ID="ddl_EDATE_HH" runat="server"></asp:DropDownList>時
                    <asp:DropDownList ID="ddl_EDATE_MM" runat="server"></asp:DropDownList>分
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">抬頭：</td>
                <td class="whitecol">
                    <asp:TextBox ID="B_TITLE" Width="80%" runat="server" placeholder="請輸入抬頭" MaxLength="1000"></asp:TextBox></td>
            </tr>
            <%--<tr>
                <td width="20%" class="bluecol_need">內容：</td>
                <td class="whitecol">
                    <asp:TextBox ID="B_CONTENT" Width="80%" runat="server" placeholder="請輸入內容"></asp:TextBox></td>
            </tr>--%>
            <tr>
                <td width="20%" class="bluecol_need">連線網頁(https)：</td>
                <td class="whitecol">
                    <asp:TextBox ID="B_URL" Width="80%" runat="server" placeholder="請輸入連線網頁" MaxLength="1000"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">提示訊息：</td>
                <td class="whitecol">
                    <asp:TextBox ID="B_ALT" Width="80%" runat="server" placeholder="請輸入提示訊息" MaxLength="1000"></asp:TextBox></td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">檔名(.jpg)：</td>
                <td class="whitecol">
                    <asp:TextBox ID="FILE_NAME" Width="50%" runat="server" placeholder="請輸入檔名" MaxLength="1000"></asp:TextBox>
                    <asp:Label ID="LabFNAME_MSG" runat="server" ForeColor="Red"></asp:Label>
                    <asp:HiddenField ID="Hid_ORG_FILE_NAME" runat="server" />
                </td>
            </tr>
            <tr>
                <td class="bluecol">上傳圖片,檔名,位置 </td>
                <td class="whitecol">
                    <input id="File1" type="file" name="File1" runat="server" accept=".jpg" size="66" />
                    <asp:Button ID="ButUpload1" runat="server" Text="確定上傳" CausesValidation="False" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">啟用：</td>
                <td class="whitecol">
                    <asp:CheckBox ID="cb_ISUSED" runat="server" />
                </td>
            </tr>
            <tr>
                <td width="20%" class="bluecol_need">順序：</td>
                <td class="whitecol">
                    <asp:TextBox ID="TXT_SEQ" Width="20%" runat="server" placeholder="請輸入提示訊息" MaxLength="10"></asp:TextBox></td>
            </tr>
            <tr>
                <td class="whitecol" align="center" colspan="2">
                    <asp:Button ID="btnSave1" runat="server" Text="儲存" CssClass="asp_button_S" AuthType="SAVE"></asp:Button>
                    &nbsp;<asp:Button ID="btnCancle1" runat="server" Text="取消" CssClass="asp_button_S" AuthType="CANCLE"></asp:Button>&nbsp;<asp:Button ID="btnBack1" runat="server" Text="回上頁" CssClass="asp_button_S" AuthType="BACK"></asp:Button>&nbsp;
                </td>
            </tr>
        </table>

        <asp:HiddenField ID="Hid_BANNERID" runat="server" />
        <asp:HiddenField ID="Hid_B_CONTENT" runat="server" />
        <asp:HiddenField ID="Hid_TYPEID" runat="server" />
    </form>
</body>
</html>
