<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_035.aspx.vb" Inherits="WDAIIP.SYS_03_035" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
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
        /* 依使用者瀏覽器環境設定物件顯示語法 */
        function ctrlShowObj(obj) {
            if (navigator.userAgent.toLowerCase().indexOf('msie') > 0) {
                //obj.style.display = 'inline';
                obj.style.display = '';
            } else {
                obj.style.display = 'table-row';
            }
        }

        /* 取得 RadioButtonList 值 */
        function getRBLValue(strObjID) {
            var i = 0;
            var strRtn = '';
            while (document.getElementById(strObjID + '_' + i)) {
                if (document.getElementById(strObjID + '_' + i).checked) {
                    strRtn = document.getElementById(strObjID + '_' + i).value;
                    break;
                }
                i += 1;
            }
            return strRtn;
        }

        //週期項目
        function chkSchScope() {
            var trYear = document.getElementById("trYear");
            var trMonth = document.getElementById("trMonth");
            var trDay = document.getElementById("trDay");
            var strQueType = "";
            var strMsg = "";

            strQueType = getRBLValue("rblScope"); //取得 RadioButtonList 值 

            trYear.style.display = "none";
            trMonth.style.display = "none";
            trDay.style.display = "none";
            if (strQueType == "Y") {
                ctrlShowObj(trYear); //不同瀏覽器的表格，顯示 inline，table-row
            }
            if (strQueType == "M") { //月
                ctrlShowObj(trMonth); //不同瀏覽器的表格，顯示 inline，table-row
            }
            if (strQueType == "D") {
                ctrlShowObj(trDay); //不同瀏覽器的表格，顯示 inline，table-row
            }
        }

    </script>
</head>
<body>
    <form id="form1" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">
				    首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;功能使用盤點
                    </asp:Label>
                </td>
            </tr>
        </table>

        <table id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <%-- <tr>
            <td>
                <table class="font" width="100%">
                    <tr>
                        <td class="font">
                            首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;<font color="#990000">功能使用盤點</font>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>--%>
            <tr>
                <td>
                    <table id="SchTable" class="table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width:20%">功能名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtFunName" runat="server" MaxLength="50" Columns="50" Width="60%"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">日期範圍
                            </td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblScope" runat="server" RepeatDirection="Horizontal" CssClass="font">
                                    <asp:ListItem Value="Y">年</asp:ListItem>
                                    <asp:ListItem Value="M">月</asp:ListItem>
                                    <asp:ListItem Value="D">日</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="trYear" runat="server" style="display: none">
                            <td class="bluecol" width="10%">年度
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddlY1" runat="server">
                                </asp:DropDownList>
                                ～
                            <asp:DropDownList ID="ddlY2" runat="server">
                            </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="trMonth" runat="server" style="display: none">
                            <td class="bluecol" width="10%">月份
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddlMY1" runat="server">
                                </asp:DropDownList>
                                <asp:DropDownList ID="ddlMM1" runat="server">
                                </asp:DropDownList>
                                ～
                            <asp:DropDownList ID="ddlMY2" runat="server">
                            </asp:DropDownList>
                                <asp:DropDownList ID="ddlMM2" runat="server">
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="trDay" runat="server" style="display: none">
                            <td class="bluecol" width="10%">日期
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="MDATE1" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= MDATE1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                                ～
                            <asp:TextBox ID="MDATE2" Width="15%" onfocus="this.blur()" runat="server"></asp:TextBox>
                                <img style="cursor: pointer" onclick="javascript:show_calendar('<%= MDATE2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" colspan="2">
                                <p align="center">
                                    <%--<asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label><asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="23px">10</asp:TextBox>--%>
                                    <asp:Button ID="btnSearch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <%-- <p align="center">
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
                </p>--%>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="DataTable1" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td align="center">
                                <div id="Div1" runat="server">
                                    <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" PageSize="20" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="編號">
                                                <ItemStyle HorizontalAlign="Center" Width="6%"></ItemStyle>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能路徑">
                                                <ItemStyle HorizontalAlign="Center" Width="40%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labFunPath" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能名稱">
                                                <ItemStyle HorizontalAlign="Center" Width="18%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labFunName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="點擊次數">
                                                <ItemStyle HorizontalAlign="Center" Width="18%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labCount" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <ItemStyle HorizontalAlign="Center" Width="18%" Font-Size="Small"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="btnListData1" runat="server" Text="明細查詢" CommandName="ListData1" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Font-Size="Medium" Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                        <%--<PagerStyle HorizontalAlign="Right" Mode="NumericPages"></PagerStyle>--%>
                                    </asp:DataGrid>
                                </div>
                                <asp:Label ID="lab_Msg1" runat="server" ForeColor="Red" Visible="False">查無資料!</asp:Label>
                                <%--<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>--%>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="btnExport1" runat="server" Text="匯出統計" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="btnExport2" runat="server" Text="匯出明細" CssClass="asp_button_M"></asp:Button>
                                <asp:Button ID="btnBack1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="DataTable2" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" style="width:20%">功能路徑
                            </td>
                            <td class="whitecol" style="width:30%">
                                <asp:Label ID="lFunPath" runat="server"></asp:Label>
                            </td>
                            <td class="bluecol" style="width:20%">功能名稱
                            </td>
                            <td class="whitecol" style="width:30%">
                                <asp:Label ID="lFunName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4">
                                <div id="Div2" runat="server">
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" AllowPaging="True" AutoGenerateColumns="False" CssClass="font" PageSize="20" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="編號" HeaderStyle-Width="5%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能路徑" HeaderStyle-Width="10%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labFunPath" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能名稱" HeaderStyle-Width="10%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labFunName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="計畫年度" HeaderStyle-Width="9%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labYears" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="所屬分署" HeaderStyle-Width="10%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labDistName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="計畫名稱" HeaderStyle-Width="9%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labTPlanName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="使用計畫" HeaderStyle-Width="10%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labPlanName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="單位名稱" HeaderStyle-Width="10%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labOrgName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="使用者" HeaderStyle-Width="9%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labAcctName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="帳號" HeaderStyle-Width="9%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labAccount" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="使用日期" HeaderStyle-Width="9%">
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labMDATE" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Font-Size="Medium" Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                    </asp:DataGrid>
                                </div>
                                <asp:Label ID="lab_Msg2" runat="server" ForeColor="Red" Visible="False">查無資料!</asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4" class="whitecol">
                                <p>
                                    <asp:Button ID="btnBack2" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <input id="hidfunid" type="hidden" runat="server" name="hidfunid" />
    </form>
</body>
</html>
