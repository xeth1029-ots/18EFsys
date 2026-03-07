<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_05_036.aspx.vb" Inherits="WDAIIP.SD_05_036" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>重大災害受災地區範圍</title>
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function getDISASTER2() {
            var hid_ADID = document.getElementById('hid_ADID');
            var Label1 = document.getElementById('Label1');
            var hid_ZIPCODE = document.getElementById('hid_ZIPCODE');

            var vrADID = hid_ADID.value;
            if (hid_ADID.value == "") { vrADID = "NEW1"; }
            Label1.innerHTML = '';
            //window.open('SD_05_036c.aspx?ADID=' + vrADID + '&ZIPCODE=' + hid_ZIPCODE.value, '', 'width=450,height=400,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
            window.open('SD_05_036c.aspx?ADID=' + vrADID, '', 'width=450,height=400,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table id="FrameTable" class="font" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="Table2" class="font" cellspacing="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">重大災害受災地區範圍</font>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table id="tbSearch1" runat="server" class="table_nw" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="20%">重大災害名稱 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="schCNAME" runat="server" MaxLength="50"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="4" align="center">
                                <p style="margin-bottom: 3px; margin-top: 3px" align="center">&nbsp;&nbsp;</p>
                                <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                                <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                                &nbsp;<asp:Button ID="btnSearch1" runat="server" Text="查詢" CssClass="asp_button_M" />
                                &nbsp;
							<asp:Button ID="btnInsert1" runat="server" Text="新增" CssClass="asp_button_M" />
                            </td>
                        </tr>
                    </table>
                    <p style="margin-bottom: 3px; margin-top: 3px" align="center">
                        <asp:Label ID="msg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </p>
                    <table class="font" id="tbDataGrid1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="center">
                                <asp:DataGrid ID="DataGrid1" runat="server" PagerStyle-Visible="False" AutoGenerateColumns="False" AllowPaging="true" AllowSorting="true" Width="100%">
                                    <AlternatingItemStyle BackColor="#EEEEEE" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="CNAME" HeaderText="重大災害名稱"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="BEGDATE" HeaderText="起始日期"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ENDDATE" HeaderText="結束日期"></asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" Width="150" Wrap="false"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lbtUpdate1" runat="server" Text="修改" CommandName="update1" CssClass="asp_button_M"></asp:LinkButton>&nbsp;
											<asp:LinkButton ID="lbtDelete1" runat="server" Text="刪除" CommandName="delete1" CssClass="asp_button_M"></asp:LinkButton>&nbsp;
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                            </td>
                        </tr>
                    </table>
                    <table id="tbDetail1" runat="server" class="font" cellspacing="1" cellpadding="1" width="100%">
                        <tr>
                            <td class="bluecol_need" width="20%">重大災害名稱 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="CNAME" runat="server" MaxLength="50"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" width="20%">起始日期 </td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="BEGDATE" runat="server" MaxLength="10"></asp:TextBox>
                                <img style="cursor: pointer;" onclick="javascript:show_calendar('BEGDATE','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="24" height="24" />
                            </td>
                            <td class="bluecol_need" width="20%">結束日期 </td>
                            <td class="whitecol" width="30%">
                                <asp:TextBox ID="ENDDATE" runat="server" MaxLength="10"></asp:TextBox>
                                <img style="cursor: pointer;" onclick="javascript:show_calendar('ENDDATE','','','CY/MM/DD');" alt="" align="top" src="../../images/show-calendar.gif" width="24" height="24" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">系統告警訊息 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="ALARMMSG1" runat="server" MaxLength="500" Width="500px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">備註 </td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="MEMO1" runat="server" MaxLength="500" Width="500px"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">重大災害受災地區 </td>
                            <td colspan="3" class="whitecol">
                                <input id="Button3" disabled="disabled" onclick="getDISASTER2();" type="button" value="選擇災害受災地區" runat="server" class="button_b_L">
                                <asp:Button ID="Button4" runat="server" Text="查詢完整受災地區" />
                                <asp:Label ID="Label1" runat="server" CssClass="font"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">使用功能 </td>
                            <td colspan="3" class="whitecol">
                                <asp:CheckBox ID="CB1_FUNC1" runat="server" Text="主要參訓身分別「經公告之重大災害受災者」子選項" Checked="True" /><br />
                                <asp:CheckBox ID="CB1_FUNC2" runat="server" Text="提醒承辦注意(限定計畫)系統告警訊息必填" Checked="True" />
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4">
                                <p style="margin-bottom: 3px; margin-top: 3px" align="center">&nbsp;&nbsp;</p>
                                <asp:Button ID="btnSaveData1" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                &nbsp;
							<asp:Button ID="btnBackup1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <%--<tr> <td class="whitecol" colspan="4"><font style="color: #009900;">*記得維護。</font> </td> </tr>--%>
                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="hid_ADID" runat="server" />
        <input id="hid_ZIPCODE" runat="server" type="hidden" />
        <input id="hid_sessName1" runat="server" type="hidden" />
    </form>
</body>
</html>
