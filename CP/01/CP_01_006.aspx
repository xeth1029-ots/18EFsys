<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="CP_01_006.aspx.vb" Inherits="WDAIIP.CP_01_006" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>不預告實地抽訪紀錄表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <%--<script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>--%>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //決定date-picker元件使用的是西元年or民國年，by:20181018
        var bl_rocYear = "<%=ConfigurationManager.AppSettings("REPLACE2ROC_YEARS") %>";
        var scriptTag = document.createElement('script');
        var jsPath = (bl_rocYear == "Y" ? "../../js/date-picker2.js" : "../../js/date-picker.js");
        scriptTag.src = jsPath;
        document.head.appendChild(scriptTag);

        function search() {
            //var msg = '';
            //var start_date = document.getElementById("start_date");
            //var end_date = document.getElementById("end_date");
            //if (start_date.value == '') msg += '請選擇起始日期\n';
            //if (end_date.value == '') msg += '請選擇終至日期\n';
            /*
            if (msg != '') {
            alert(msg);
            return false;
            }
            */
        }

        function CheckAdd() {
            var OCIDValue1 = document.getElementById("OCIDValue1");
            if (OCIDValue1.value == '') {
                alert('請選擇職類班別!')
                return false;
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;不預告實地抽訪紀錄表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="FrameTable3" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td class="bluecol" width="20%">機構</td>
                <td class="whitecol">
                    <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                    <input id="Button5" type="button" value="..." name="Button5" runat="server" class="asp_button_Mini">
                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server">
                    <span id="HistoryList2" style="position: absolute; display: none">
                        <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol">職類/班別 </td>
                <td class="whitecol">
                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                    <input id="Button6" onclick="javascript: openClass('../CP_01_ch.aspx?RID=' + document.form1.RIDValue.value);" type="button" value="..." name="Button6" runat="server" class="asp_button_Mini">
                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                    <span id="HistoryList" style="position: absolute; display: none; left: 270px">
                        <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                    </span></td>
            </tr>
            <tr>
                <td class="bluecol_need"><font>訪查日期</font> </td>
                <td class="whitecol">
                    <span runat="server">
                        <asp:TextBox ID="start_date" runat="server" Width="18%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= start_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        ～<asp:TextBox ID="end_date" runat="server" Width="18%"></asp:TextBox>
                        <img style="cursor: pointer" onclick="javascript:show_calendar('<%= end_date.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                    </span>
                </td>
            </tr>
            <tr>
                <td class="bluecol_need"><font>抽訪狀態</font> </td>
                <td class="whitecol">
                    <asp:RadioButtonList ID="VisitItem" runat="server" RepeatDirection="Horizontal" CssClass="font">
                        <asp:ListItem Value="1" Selected="True">未抽訪</asp:ListItem>
                        <asp:ListItem Value="2">已抽訪</asp:ListItem>
                        <asp:ListItem Value="3">全部</asp:ListItem>
                    </asp:RadioButtonList>
                    &nbsp; (選取未抽訪時，訪查日期選項不列入查詢項目) </td>
            </tr>
            <tr>
                <td colspan="2" class="whitecol">
                    <p align="center">
                        <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue" DESIGNTIMEDRAGDROP="30">顯示列數</asp:Label>
                        <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>
                        <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button2" runat="server" Text="新增" CssClass="asp_button_M"></asp:Button>
                        <asp:Button ID="Button10" runat="server" Text="列印空白表單" CssClass="asp_Export_M"></asp:Button>
                    </p>
                </td>
            </tr>
            <tr>
                <td colspan="2" class="whitecol">
                    <div style="width: 100%" align="center">
                        <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="2" class="whitecol">
                    <table id="DataGridTable" width="100%" runat="server">
                        <tr>
                            <td class="whitecol" align="left">
                                <asp:Label ID="labmsg1" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn HeaderText="序號">
                                            <HeaderStyle HorizontalAlign="Center" Width="5%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="OrgName" HeaderText="訓練機構"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="ClassCName" HeaderText="班別名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDATE_TW" HeaderText="開訓日期"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDATE_TW" HeaderText="結訓日期"></asp:BoundColumn>
                                        <asp:BoundColumn DataField="APPLYDATE_TW" HeaderText="訪查日期"></asp:BoundColumn>
                                        <%--<asp:TemplateColumn HeaderText="訪查日期">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%" VerticalAlign="Middle"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="LabApplyDate" runat="server"></asp:Label></ItemTemplate>
                                        </asp:TemplateColumn>--%>
                                        <asp:BoundColumn HeaderText="抽訪結果">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%" VerticalAlign="Middle"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center" VerticalAlign="Middle"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="原因及後續追蹤">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%" VerticalAlign="Middle"></HeaderStyle>
                                            <ItemTemplate>
                                                <textarea id="Reason" style="width: 100%; height: 70px" name="VerReason" rows="3" cols="15" runat="server"></textarea>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>

                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%" VerticalAlign="Middle"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Button ID="Button3e" runat="server" Text="修改" CommandName="edit" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button3v" runat="server" Text="查詢" CommandName="view" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button4" runat="server" Text="刪除" CommandName="del" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="BtnAddStd" runat="server" Text="抽訪學員紀錄" CommandName="AddStd" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button8" runat="server" Text="列印" CommandName="prt1" CssClass="asp_Export_M"></asp:Button>
                                                <asp:Button ID="Button7" runat="server" Text="新增" CommandName="Add" CssClass="asp_button_M"></asp:Button>
                                                <asp:Button ID="Button11" runat="server" Text="列印空白表單" CommandName="prt2" CssClass="asp_Export_M"></asp:Button>
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

        <%--<table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0"><tbody></tbody></table>--%>
    </form>
</body>
</html>
