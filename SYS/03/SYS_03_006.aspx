<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_006.aspx.vb" Inherits="WDAIIP.SYS_03_006" %>


<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>已結訓班級使用授權設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
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
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;已結訓班級使用授權設定</asp:Label>
                </td>
            </tr>
        </table>
        <input id="check_del" type="hidden" name="check_del" runat="server" />
        <input id="check_mod" type="hidden" name="check_mod" runat="server" />
        <input id="check_add" type="hidden" name="check_add" runat="server" />
        <input id="check_Sech" type="hidden" name="check_Sech" runat="server" />
        <%--    
    <asp:TextBox ID="IntStr" runat="server" Visible="False" Columns="1"></asp:TextBox>
    <asp:TextBox ID="EditStr" runat="server" Visible="False" Columns="1"></asp:TextBox>
    <asp:TextBox ID="DelStr" runat="server" Visible="False" Columns="1"></asp:TextBox>
    <asp:TextBox ID="Cnt" runat="server" Visible="False" Columns="1"></asp:TextBox>--%>
        <table class="font" id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="table2" cellspacing="1" cellpadding="1" width="100%" border="0">
					<tr>
						<td>
							<asp:Label ID="TitleLab1" runat="server"></asp:Label>
							<asp:Label ID="TitleLab2" runat="server">
				            首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt; <font color="#990000">已結訓班級使用授權設定</font>
							</asp:Label>
						</td>
					</tr>
				</table>--%>
                    <table class="font" id="table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <table class="table_nw" id="Searchtable" cellspacing="1" width="100%">
                                    <tr>
                                        <td class="bluecol" width="20%">&nbsp; 年度：
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="yearlist" runat="server" AutoPostBack="true">
                                            </asp:DropDownList>
                                        </td>
                                        <td class="bluecol" width="20%">&nbsp;轄區：
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="DistID" runat="server" AutoPostBack="true">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">&nbsp; 訓練計畫：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:DropDownList ID="planlist" runat="server" AutoPostBack="true">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">&nbsp; 機構名稱：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="schOrgName" runat="server" MaxLength="50" Width="40%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">&nbsp; 班級名稱：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="ClassName" runat="server" MaxLength="50" Width="90%"></asp:TextBox>
                                        </td>
                                        <td class="bluecol" width="20%">&nbsp;期別：
                                        </td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="CyclType" runat="server" Columns="5" Width="40%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">&nbsp; 開訓日期：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <span id="span01" runat="server">
                                                <asp:TextBox ID="start_date" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
										    <asp:TextBox ID="end_date" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                            </span>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="20%">班級範圍：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:RadioButtonList ID="ClassRound" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatLayout="Flow">
                                                <asp:ListItem Value="C" Selected="true">已結訓</asp:ListItem>
                                                <asp:ListItem Value="O">未結訓</asp:ListItem>
                                                <asp:ListItem Value="A">全部</asp:ListItem>
                                            </asp:RadioButtonList>
                                        </td>
                                    </tr>
                                </table>
                                <table width="100%">
                                    <tr class="whitecol">
                                        <td align="center" colspan="4">
                                            <asp:Button ID="rt_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button><br />
                                            <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                                        </td>
                                    </tr>
                                </table>
                                <table width="100%" class="table_sch" cellpadding="1" cellspacing="1">
                                    <tr id="trOrgName" runat="server">
                                        <td class="bluecol" width="20%">取得授權單位：
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:DropDownList ID="ddlOrgName" runat="server" AutoPostBack="true">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr id="Account_tr" runat="server">
                                        <td class="bluecol" width="20%">取得授權帳號：
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="Account" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                        <td class="bluecol" width="20%">補登資料原因：
                                        </td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="ReasonID" runat="server">
                                            </asp:DropDownList>
                                        </td>
                                    </tr>
                                    <tr id="Reason_tr" runat="server">
                                        <td class="bluecol" width="20%">補登資料原因簡述
                                        </td>
                                        <td class="whitecol" colspan="3">
                                            <asp:TextBox ID="Reason" runat="server" Columns="5" Width="100%"></asp:TextBox>
                                        </td>
                                    </tr>
                                    <tr id="Fun_tr" runat="server">
                                        <td class="bluecol" width="20%">開放功能 ：
                                        </td>
                                        <td class="whitecol">
                                            <asp:CheckBoxList ID="cb_SelFunID" runat="server" RepeatLayout="Flow" Width="100%">
                                            </asp:CheckBoxList>
                                        </td>
                                        <td class="bluecol" width="20%">結束日期 ：
                                        </td>
                                        <td class="whitecol">
                                            <span id="span02" runat="server">
                                                <asp:TextBox ID="EndDate" Width="40%" onfocus="this.blur()" runat="server"></asp:TextBox>
                                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= EndDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                            </span>
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
                    <asp:Label ID="Label2" runat="server" Visible="False">  已結訓班級資料：</asp:Label>移到該項目滑鼠停留會顯示開放功能。
                </td>
            </tr>
            <tr>
                <td>
                    <table class="font" id="tbSearch1" cellspacing="1" cellpadding="1" align="center" border="0" runat="server" width="100%">
                        <tr>
                            <td align="center">
                                <asp:DataGrid ID="DG_ClassInfo" runat="server" CssClass="font" AllowPaging="true" AutoGenerateColumns="False" PageSize="30" Width="100%" CellPadding="8">
                                    <AlternatingItemStyle BackColor="WhiteSmoke" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn Visible="False" HeaderText="選擇">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <input id="FunID" type="checkbox" name="FunID" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="OrgName" SortExpression="OrgName" HeaderText="訓練機構">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassID2" HeaderText="班別代碼">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Stdate" HeaderText="開訓日" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Ftdate" HeaderText="結訓日" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="trainName" HeaderText="訓練職類">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="OCID" HeaderText="OCID"></asp:BoundColumn>
                                        <asp:BoundColumn Visible="False" DataField="RightID" HeaderText="RightID"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="Name" HeaderText="已授權給">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="EndDate" HeaderText="結束日期" DataFormatString="{0:d}">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能" ItemStyle-HorizontalAlign="Center">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <ItemStyle Font-Size="Small" />
                                            <ItemTemplate>
                                                <asp:LinkButton ID="but1" runat="server" Text="新增" CommandName="Add" CssClass="linkbutton"></asp:LinkButton>
                                                <asp:LinkButton ID="but2" runat="server" Text="修改" CommandName="Upd" CssClass="linkbutton"></asp:LinkButton><br />
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
                            <td align="center">
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
