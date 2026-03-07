<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SYS_06_009.aspx.vb" Inherits="WDAIIP.SYS_06_009" %>


<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>個資安全密碼保護設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
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
        function getCheckBoxListItemsChecked() {
            var elementref = document.getElementById('cb_SelFunID');
            var checkBoxArray = elementref.getElementsByTagName('input');
            var checkedValues = 0;
            for (var i = 0; i < checkBoxArray.length; i++) {
                var checkBoxRef = checkBoxArray[i];
                if (checkBoxRef.checked == true) {
                    checkedValues += 1;
                }
            }
            return checkedValues;
        }

        function ChkData() {
            var msg = '';
            var checkedItems = getCheckBoxListItemsChecked();
            if (checkedItems == 0) {
                msg += '請選擇欲開放功能！\n';
            }
            if (document.getElementById('EndDate').value == '') {
                msg += '請選擇結束日期！';
            }
            if (msg != '') {
                alert(msg);
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <%--	<input id="check_del" type="hidden" name="check_del" runat="server" />
	<input id="check_mod" type="hidden" name="check_mod" runat="server" />
	<input id="check_add" type="hidden" name="check_add" runat="server" />
	<input id="check_Sech" type="hidden" name="check_Sech" runat="server" />
	<asp:TextBox ID="IntStr" runat="server" Visible="False" Columns="1"></asp:TextBox>
	<asp:TextBox ID="EditStr" runat="server" Visible="False" Columns="1"></asp:TextBox>
	<asp:TextBox ID="DelStr" runat="server" Visible="False" Columns="1"></asp:TextBox>
	<asp:TextBox ID="Cnt" runat="server" Visible="False" Columns="1"></asp:TextBox>--%>
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">
				            首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt; <font color="#990000">個資安全密碼保護設定</font>
                    </asp:Label>
                </td>
            </tr>
        </table>

        <table class="font" id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>

                    <table class="font" id="table3" cellspacing="1" cellpadding="1" width="740" border="0">
                        <tr>
                            <td>
                                <table class="table_nw" id="Searchtable" cellspacing="1" width="740">
                                    <tr>
                                        <td class="bluecol_need" style="width: 100px;">&nbsp; 年度：
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="true">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">&nbsp;轄區：
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="DistID" runat="server" AutoPostBack="true" Width="160px">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need">&nbsp; 訓練計畫：
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="planlist" runat="server" AutoPostBack="true" Width="420px">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                </table>
                                <table width="100%">
                                    <tr class="whitecol">
                                        <td align="center" colspan="4">
                                            <asp:Button ID="rt_search" runat="server" Text="查詢" CssClass="asp_button_S"></asp:Button>&nbsp;
										<asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                                <table width="100%" class="table_sch" cellpadding="1" cellspacing="1">
                                    <tr id="trOrgName" runat="server">
                                        <td class="bluecol_need" style="width: 100px;">&nbsp; 設定單位：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:DropDownList ID="ddlOrgName" runat="server" AutoPostBack="true">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr id="Account_tr" runat="server">
                                        <td class="bluecol_need">設定帳號：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:DropDownList ID="Account" runat="server" Width="250px">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol_need" style="width: 100px;">&nbsp; 開始日期時間 ：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="SDATE" Width="80" onfocus="this.blur()" runat="server"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= SDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                        </td>
                                        <td class="bluecol_need" style="width: 100px;">&nbsp; 結束日期時間 ：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="EDATE" Width="80" onfocus="this.blur()" runat="server"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= EDATE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                        </td>
                                    </tr>
                                    <tr id="Fun_tr" runat="server">
                                        <td class="bluecol_need">&nbsp; 開放功能 ：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:CheckBoxList ID="cb_SelFunID" runat="server" RepeatLayout="Flow">
                                            </asp:CheckBoxList>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="whitecol">
                    <%--
				<asp:Label ID="Label2" runat="server" Visible="False">  已結訓班級資料：</asp:Label>
				移到該項目滑鼠停留會顯示開放功能。
                    --%>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="DataGridTable1" border="0" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td align="center">
                                <asp:DataGrid Style="z-index: 0" ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font">
                                    <AlternatingItemStyle BackColor="White"></AlternatingItemStyle>
                                    <ItemStyle BackColor="#EBF3FE"></ItemStyle>
                                    <HeaderStyle HorizontalAlign="Center" BackColor="#96B5E3"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="編號">
                                            <HeaderStyle Width="30px"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ACCOUNT" HeaderText="使用者帳號"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="CName" HeaderText="姓名"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="WSDate" HeaderText="作業開始日期"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="WEDate" HeaderText="作業結束日期"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <ItemTemplate>
                                                <asp:LinkButton ID="but1" runat="server" Text="新增" CommandName="Add" CssClass="linkbutton"></asp:LinkButton>
                                                <asp:LinkButton ID="but2" runat="server" Text="修改" CommandName="Upd" CssClass="linkbutton"></asp:LinkButton>
                                                <asp:LinkButton ID="but3" runat="server" Text="刪除" CommandName="Del" CssClass="linkbutton"></asp:LinkButton>
                                                <asp:LinkButton ID="but4" runat="server" Text="取得" CommandName="GetData" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td style="height: 31px" align="center">
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
