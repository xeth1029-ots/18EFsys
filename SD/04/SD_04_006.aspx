<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_04_006.aspx.vb" Inherits="WDAIIP.SD_04_006" %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>老師排課狀況查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/OpenWin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
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

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?RID=' + RIDValue.value);
        }

        function check_data() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == '') {
                if (parseInt(getCheckBoxListValue('InTeach')) == 0 && parseInt(getCheckBoxListValue('OutTeach')) == 0) {
                    alert('請選擇老師');
                    return false;
                }
            }
        }

        function GetAllTeach(Mode, flag) {
            var obj = (Mode == 1) ? 'InTeach' : 'OutTeach';
            var objcount = getCheckBoxListValue(obj).length;
            for (var i = 0; i < objcount; i++) {
                document.getElementById(obj + '_' + i).checked = flag;
            }
        }

        function GetAllCourse(Mode, flag) {
            var j = 0;
            var k = 0;

            switch (Mode) {
                case 1:
                    j = 0;
                    k = 3;
                    break;
                case 2:
                    j = 4;
                    k = 7;
                    break;
                case 3:
                    j = 8;
                    k = 11;
                    break;
            }

            for (var i = j; i <= k; i++) {
                document.getElementById('ClassNum_' + i).checked = flag;
            }
        }
    </script>

</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td align="center">
                    <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;課程管理&gt;&gt;老師排課狀況查詢</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <asp:Panel ID="SearchTable" runat="server">
                        <table class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                            <tr>
                                <td class="bluecol" style="width: 20%">訓練機構
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="center" runat="server" Width="55%"></asp:TextBox>
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                                    <input id="Button8" type="button" value="..." name="Button5" runat="server" class="button_b_Mini">
                                    <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                    <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">職類/班別
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                    <asp:Button ID="clear" runat="server" Text="清除" CssClass="asp_button_S"></asp:Button>
                                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                    <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                    </span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">
                                    <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini">
                                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">老師姓名
                                </td>
                                <td class="whitecol">
                                    <div>
                                        <asp:TextBox ID="txtSchTeachCName" Width="20%" runat="server"></asp:TextBox>
                                        <asp:Button ID="btnSchTeach" class="button_b_Mini" Text="老師查詢" runat="server" />
                                        <font color="red">
                                            <asp:Label ID="Label1" runat="server" Text="(查詢該單位老師)"></asp:Label></font>
                                    </div>
                                    <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0">
                                        <tr>
                                            <td>內聘
												<asp:CheckBox ID="InSelectAll" runat="server" Text="全選"></asp:CheckBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:CheckBoxList ID="InTeach" runat="server" RepeatColumns="6" RepeatDirection="Horizontal" CssClass="font">
                                                </asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div id="divInTeachChk"></div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>外聘
												<asp:CheckBox ID="OutSelectAll" runat="server" Text="全選"></asp:CheckBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:CheckBoxList ID="OutTeach" runat="server" RepeatColumns="6" RepeatDirection="Horizontal" CssClass="font">
                                                </asp:CheckBoxList>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div id="divOutTeachChk"></div>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">日期範圍
                                </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="SDate" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= SDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">～
									<asp:TextBox ID="EDate" runat="server" Columns="10" Width="15%"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= EDate.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">節 次
                                </td>
                                <td class="whitecol">
                                    <table class="whitecol" id="Table4" cellspacing="0" cellpadding="0" width="100%" border="0">
                                        <tr>
                                            <td>
                                                <asp:CheckBox ID="CourseRound1" runat="server" Text="1-4節"></asp:CheckBox><asp:CheckBox ID="CourseRound2" runat="server" Text="5-8節"></asp:CheckBox><asp:CheckBox ID="CourseRound3" runat="server" Text="9-12節"></asp:CheckBox>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <asp:CheckBoxList ID="ClassNum" runat="server" RepeatColumns="6" RepeatDirection="Horizontal" CssClass="font" RepeatLayout="Flow">
                                                </asp:CheckBoxList>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">老師助教
                                </td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rblTeachtype" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Selected="True" Value="1">教師1</asp:ListItem>
                                        <asp:ListItem Value="2">助教1</asp:ListItem>
                                        <asp:ListItem Value="3">助教2</asp:ListItem>
                                        <%--<asp:ListItem Value="4">教師2</asp:ListItem>
										<asp:ListItem Value="5">教師3</asp:ListItem>--%>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="btnPrint" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table class="font" id="DataGridTable" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AllowPaging="True" PageSize="15" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="TeachCName" HeaderText="老師姓名">
                                            <HeaderStyle Width="10%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SchoolDate" HeaderText="日期" DataFormatString="{0:d}">
                                            <HeaderStyle Width="10%" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班級名稱">
                                            <HeaderStyle Width="35%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CourseName" HeaderText="課程名稱">
                                            <HeaderStyle Width="35%" />
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassNum" HeaderText="節次">
                                            <HeaderStyle Width="10%" />
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundColumn>
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
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="Button2" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
