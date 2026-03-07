<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_03_002_ver.aspx.vb" Inherits="WDAIIP.SD_03_002_ver" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員資料審核</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function SetOneOCID() {
            document.getElementById('Button4').click();
        }

        function search() {
            if (document.form1.OCIDValue1.value == '') {
                alert('請選擇職類班別!')
                return false;
            }
        }

        function choose_class(num) {
            var RID = document.form1.RIDValue.value;
            document.form1.TMID1.value = '';
            document.form1.TMIDValue1.value = '';
            document.form1.OCID1.value = '';
            document.form1.OCIDValue1.value = '';
            //document.form1.Button2.disabled=true;
            //document.getElementById('ImportTable').style.display='none';//匯入學員名冊
            //document.getElementById('DataGridTable').style.display='none';//學員資料
            //document.getElementById('msg').innerHTML='';	
            if (document.getElementById('OCID1').value == '')
            { document.getElementById('Button4').click(); }
            openClass('../02/SD_02_ch.aspx?RWClass=1&RID=' + RID);
        }

        function but_edit(ocid, id) {
            location.href = 'TC_01_004_add.aspx?ocid=' + ocid + '&ProcessType=Update&ID=' + id;
        }

        function but_del(ocid, PlanID, ComIDNO, SeqNO, Years, is_parent, id) {
            if (is_parent) {
                alert("此班級檔尚有與班級學員檔或排課檔或已有報名資料參照,不可刪除!!");
                return;
            }
            if (window.confirm("此動作會刪除班別資料，是否確定刪除?"))
                location.href = 'TC_01_004_del.aspx?ocid=' + ocid + '&PlanID=' + PlanID + '&ComIDNO=' + ComIDNO + '&SeqNO=' + SeqNO + '&Years=' + Years + '&ID=' + id;
        }

        function check_value() {
            /*if (document.form1.TPeriod_List.value=='03'){
			alert("目前暫時不能選擇此項目");
			document.form1.TPeriod_List.value='';
			}*/
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="font">
                    <%--首頁&gt;&gt;學員動態管理&gt;&gt;報到&gt;&gt;<FONT color="#990000">學員資料審核</FONT>--%>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;學員資料查詢</asp:Label>
                    <%--
                    <input id="check_del" style="width: 56px; height: 22px" type="hidden" size="4" name="check_del" runat="server" />
				    <input id="check_mod" style="width: 56px; height: 22px" type="hidden" size="4" name="check_mod" runat="server" />
				    <input id="check_add" style="width: 56px; height: 22px" type="hidden" size="4" name="check_add" runat="server" />
                    --%>
                </td>
            </tr>
        </table>
        <%--<asp:Label ID="Label1" runat="server" CssClass="font">在[學員資料維護]功能裡按[學員資料確認]鍵後,才會出現未審核的學員資料</asp:Label><br>--%>
        <table class="table_nw" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td class="bluecol_need" width="20%">訓練機構 </td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="Org" type="button" value="..." name="Org" runat="server" class="asp_button_Mini" />
                    <input id="RIDValue" style="width: 32px;" type="hidden" name="RIDValue" runat="server" />
                    <asp:Button ID="Button4" Style="display: none" runat="server"></asp:Button>
                    <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                    <span id="HistoryList2" style="display: none; position: absolute" onclick="GETvalue()">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">職類/班級 </td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <input id="Button5" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini" />
                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                    <span id="HistoryList" style="display: none; left: 30%; position: absolute">
                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <%--
				<TR>
					<TD align="left" bgColor="#cccc66">&nbsp;&nbsp;&nbsp;&nbsp;班級名稱</TD>
					<TD><asp:textbox id="TB_ClassName" runat="server"></asp:textbox></TD>
					<TD bgColor="#cccc66"><FONT face="新細明體">&nbsp;&nbsp;&nbsp;&nbsp;</FONT>期別</TD>
					<TD><FONT face="新細明體"><asp:textbox id="TB_cycltype" runat="server" Width="40px" Columns="5"></asp:textbox></FONT></TD>
				</TR>
            --%>
            <tr>
                <td id="td5" class="bluecol" runat="server" width="20%">開訓日期 </td>
                <td class="whitecol" width="80%">
                    <asp:TextBox ID="start_date" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span> ～
                    <asp:TextBox ID="end_date" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                    <span runat="server">
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                </td>
            </tr>
            <tr>
                <td class="bluecol" width="20%">審核狀態 </td>
                <td class="whitecol" colspan="3" width="80%">
                    <asp:RadioButtonList ID="NotOpen" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
                        <%---<asp:ListItem Value="N" Selected="True">未審核</asp:ListItem>--%>
                        <asp:ListItem Value="Y" Selected="True">已審核</asp:ListItem>
                        <asp:ListItem Value="R">退件修正</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
            <tr>
                <td class="whitecol" colspan="4" align="center" width="100%">
                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                    <asp:Button ID="bt_search" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button><br />
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
        <!--<TABLE id="Table1" cellSpacing="1" cellPadding="1" width="600" border="0"></TABLE>-->
        <asp:Panel ID="Panel" runat="server" Width="100%" Visible="False">
            <table class="font" id="search_tbl" cellspacing="0" cellpadding="0" width="100%" border="1" runat="server"></table>
            <asp:DataGrid ID="DG_ClassInfo" runat="server" CssClass="font" Width="100%" Visible="False" AllowSorting="True" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                <AlternatingItemStyle BackColor="#F5F5F5" />
                <HeaderStyle CssClass="head_navy" />
                <Columns>
                    <asp:BoundColumn HeaderText="序號">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="ORGNAME" SortExpression="OrgName" HeaderText="訓練機構">
                        <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="OCLASSID" HeaderText="班別代碼">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn HeaderText="開結訓日">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                    </asp:BoundColumn>
                    <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班別名稱">
                        <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                    </asp:BoundColumn>
                    <%--  <asp:BoundColumn Visible="False" DataField="TPropertyID" HeaderText="訓練性質">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                    </asp:BoundColumn>--%>
                    <asp:TemplateColumn HeaderText="功能">
                        <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                        <ItemTemplate>
                            <asp:LinkButton ID="edit_but" runat="server" Text="學員資料審核" CommandName="edit" CssClass="linkbutton"></asp:LinkButton>
                            <asp:LinkButton ID="view_but" runat="server" Text="學員資料查詢" CommandName="view" CssClass="linkbutton"></asp:LinkButton>
                        </ItemTemplate>
                    </asp:TemplateColumn>
                </Columns>
                <PagerStyle Visible="False"></PagerStyle>
            </asp:DataGrid>
            <div align="center">
                <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
            </div>
        </asp:Panel>
        <table class="font" id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5" />
                        <HeaderStyle CssClass="head_navy" />
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="ORGNAME" SortExpression="OrgName" HeaderText="訓練機構">
                                <HeaderStyle HorizontalAlign="Center" Width="30%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="OCLASSID" HeaderText="班別代碼">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn HeaderText="開結訓日">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班別名稱">
                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                            </asp:BoundColumn>
                            <%--<asp:BoundColumn Visible="False" DataField="TPropertyID" HeaderText="訓練性質">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                            </asp:BoundColumn>--%>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center" Font-Size="Small" />
                                <ItemTemplate>
                                    <asp:LinkButton ID="return_btn" runat="server" Text="審核還原" CommandName="edit2" CssClass="linkbutton"></asp:LinkButton>
                                    <asp:LinkButton ID="view_btn" runat="server" Text="學員資料查詢" CommandName="view2" CssClass="linkbutton"></asp:LinkButton>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <uc1:PageControler ID="PageControler2" runat="server"></uc1:PageControler>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
