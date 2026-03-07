<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_014.aspx.vb" Inherits="WDAIIP.SD_04_014" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>外聘師資管理</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        //function search(){
        //	var msg=''
        //if(document.form1.DropDownList1.selectedIndex==0) msg+='請選擇師資別\n';
        //if(document.form1.DropDownList2.selectedIndex==0) msg+='請選擇任職狀況\n';
        //if(document.form1.DropDownList3.selectedIndex==0) msg+='請選擇職稱\n';
        //if(document.form1.DropDownList4.selectedIndex==0) msg+='請選擇內外聘\n';
        //	if(document.getElementById('IDate1').value =='') {msg+='請輸入聘約起日!!\n';}
        //	if(document.getElementById('IDate2').value =='') {msg+='請輸入聘約迄日!!\n';}
        //	if ((Date.parse(IDate1)).valueOf() >= (Date.parse(IDate2)).valueOf())
        //　{msg += '聘約迄日不能大於聘約起日或等於聘約迄日\n';}
        //	if(document.getElementById('IDate1').value >= document.getElementById('IDate2').value)
        //	{
        //	msg+='聘約迄日不能大於聘約起日或等於聘約迄日\n';
        //	{

        //	if (msg!=''){
        //	alert(msg);
        //	return false;
        //}
        //}

        //清除主要職類
        function ClearCarrer() {
            document.getElementById('TB_career_id').value = '';
            document.getElementById('trainValue').value = '';
            //document.getElementById('jobValue').value='';
        }

        function ShowFrame() {
            document.getElementById('FrameObj').height = document.getElementById('HistoryRID').rows.length * 20;
            document.getElementById('FrameObj').style.display = document.getElementById('HistoryList2').style.display;
        }

        function Get_Teah(fieldname, hidden, TMID) {
            wopen('../../common/TechID.aspx?RID=' + document.form1.RIDValue.value + '&ValueField=' + hidden + '&TextField=' + fieldname + '&TMID=' + TMID + '&FuntionID=SD_14_014' + '&Butn=Search1', 'LessonTeah1', 350, 500, 1);
        }

        //	function closeDiv()	{ 
        //		document.getElementById('eMeng').style.visibility='hidden'; 
        //	}			
    </script>
    <style type="text/css">
        .auto-style1 { font-size: 14px; color: Black; line-height: 26px; background-color: #FFFFFF; padding: 4px; width: 293px; }
        .auto-style2 { border-style: none; border-color: inherit; border-width: 0px; background-color: #89A5D8; border-collapse: separate; border-spacing: 1px; width: 770px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <%--<asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;課程管理&gt;&gt;外聘師資管理</asp:Label>--%>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;教師資料管理&gt;&gt;外聘師資管理</asp:Label>
                </td>
            </tr>
        </table>
        <asp:Panel ID="Tsearch" runat="server">
            <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                <tr>
                    <td class="bluecol" style="width: 20%">訓練機構</td>
                    <td colspan="3" class="whitecol">
                        <asp:TextBox ID="center" runat="server" Width="40%" onfocus="this.blur()"></asp:TextBox>
                        <input id="Button5" type="button" value="..." name="Button5" runat="server" class="button_b_Mini">
                        <span id="HistoryList2" style="display: none; z-index: 1; position: absolute">
                            <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                        </span>
                        <iframe id="FrameObj" style="display: none; position: absolute" frameborder="0" width="310" scrolling="no" height="52"></iframe>
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">師資別</td>
                    <td colspan="3" class="whitecol">
                        <asp:DropDownList ID="KindID" runat="server"></asp:DropDownList></td>
                </tr>
                <tr>
                    <td class="bluecol" style="width: 20%">講師姓名</td>
                    <td class="whitecol" style="width: 30%">
                        <asp:TextBox ID="TeachCName" runat="server" Width="40%"></asp:TextBox></td>
                    <td class="bluecol" style="width: 20%">身分證號碼</td>
                    <td class="whitecol" style="width: 30%">
                        <asp:TextBox ID="IDNO" runat="server" Width="40%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol">講師代碼</td>
                    <td class="whitecol">
                        <asp:TextBox ID="TeacherID" runat="server" Width="40%"></asp:TextBox></td>
                    <td class="bluecol">主要職類</td>
                    <td class="whitecol">
                        <asp:TextBox ID="TB_career_id" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                        <input onclick="openTrain2(document.getElementById('trainValue').value);" type="button" value="..." class="button_b_Mini"><input id="btn_clear" onclick="    ClearCarrer();" type="button" value="清除" name="Button1" class="button_b_S" />
                        <input id="trainValue" style="width: 48px; height: 22px" type="hidden" size="2" name="trainValue" runat="server">
                        <input id="jobValue" type="hidden" name="jobValue" runat="server">&nbsp;
                    </td>
                </tr>
                <tr>
                    <td class="bluecol">職稱</td>
                    <td class="whitecol">
                        <asp:DropDownList ID="IVID" runat="server"></asp:DropDownList></td>
                    <td class="bluecol">聘約期限</td>
                    <td class="whitecol">
                        <asp:TextBox ID="start_date" runat="server" Width="40%"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                        <asp:TextBox ID="end_date" runat="server" Width="40%"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    </td>
                </tr>
            </table>
            <table width="100%">
                <tr>
                    <td align="center" colspan="4" class="whitecol">
                        <asp:Button ID="Search1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Add2" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <asp:Panel ID="TInsert" runat="server">
            <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                <tr id="Insert_TR" runat="server">
                    <td class="bluecol" style="width: 20%">訓練單位</td>
                    <td class="whitecol">
                        <asp:TextBox ID="IOrg" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    </td>
                </tr>
                <tr id="Insert_TR2" runat="server">
                    <td class="bluecol_need">講師姓名</td>
                    <td class="whitecol">
                        <asp:TextBox ID="IteachName" runat="server" ToolTip="點選兩下可以跳出視窗選擇教師" Columns="20" Width="30%" ReadOnly="true"></asp:TextBox>
                        <asp:TextBox ID="IteachName2" runat="server" onfocus="this.blur()" Columns="20" Width="30%"></asp:TextBox>
                        <input id="TeahValue" type="hidden" name="TeahValue" runat="server">
                        <font color="#FF5566">&nbsp;(點選兩下可以跳出視窗選擇教師)</font>
                    </td>
                </tr>
                <tr id="Insert_TR3" runat="server">
                    <td class="bluecol">職類</td>
                    <td class="whitecol">
                        <asp:TextBox ID="TMID" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox></td>
                </tr>
                <tr>
                    <td class="bluecol_need">聘約期限</td>
                    <td class="whitecol">
                        <asp:TextBox ID="IDate1" runat="server" Width="15%"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= IDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span> ~
                        <asp:TextBox ID="IDate2" runat="server" Width="15%"></asp:TextBox>
                        <span runat="server">
                            <img style="cursor: pointer" onclick="javascript:show_calendar('<%= IDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30"></span>
                    </td>
                </tr>
            </table>
            <table width="100%">
                <tr>
                    <td align="center" class="whitecol">
                        <asp:Button ID="Save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="ReFist" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                    </td>
                </tr>
            </table>
        </asp:Panel>
        <table id="TableDataGrid1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td align="center">
                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" CellPadding="8">
                        <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                        <HeaderStyle CssClass="head_navy"></HeaderStyle>
                        <Columns>
                            <asp:BoundColumn HeaderText="序號">
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TeacherID" HeaderText="講師代碼">
                                <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="TeachCName" HeaderText="講師名稱">
                                <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="IDNO" HeaderText="身分證">
                                <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:BoundColumn DataField="KindID" HeaderText="師資別">
                                <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="職類">
                                <HeaderStyle HorizontalAlign="Center" VerticalAlign="Middle" Width="15%"></HeaderStyle>
                                <ItemTemplate>
                                    <asp:Label ID="LTMID" runat="server"></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                            <asp:BoundColumn DataField="TEDate" HeaderText="聘約起迄日">
                                <HeaderStyle HorizontalAlign="Center" Width="15%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                            </asp:BoundColumn>
                            <asp:TemplateColumn HeaderText="功能">
                                <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                <ItemTemplate>
                                    <asp:Button ID="Edit" runat="server" Text="修改" CommandName="Edit" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="Del" runat="server" Text="刪除" CommandName="Del" CssClass="asp_button_M"></asp:Button>
                                </ItemTemplate>
                            </asp:TemplateColumn>
                        </Columns>
                        <PagerStyle Visible="False"></PagerStyle>
                    </asp:DataGrid>
                </td>
            </tr>
            <tr>
                <td>
                    <p align="center">
                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                    </p>
                </td>
            </tr>
        </table>
        <table width="100%">
            <tr>
                <td align="center">
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></td>
            </tr>
        </table>
    </form>
</body>
</html>
