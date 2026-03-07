<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_036.aspx.vb" Inherits="WDAIIP.SYS_03_036" %>

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
            //
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
			            首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;功能使用查詢
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
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">
							首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;<font color="#990000">功能使用查詢</font>
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>--%>
            <tr>
                <td>
                    <table id="SchTable" class="table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol_need" style="width: 20%">功能名稱
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtFunName" runat="server" MaxLength="50" Columns="50" Width="60%"></asp:TextBox>
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
                                                <ItemStyle HorizontalAlign="Center" Width="5%"></ItemStyle>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能路徑">
                                                <ItemStyle HorizontalAlign="Center" Width="45%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labFunPath" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能名稱">
                                                <ItemStyle HorizontalAlign="Center" Width="30%"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labFunName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="功能">
                                                <ItemStyle HorizontalAlign="Center" Width="20%" Font-Size="Small"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="btnListData1" runat="server" Text="群組查詢" CommandName="ListData1" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Font-Size="Medium" Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                        <%--<PagerStyle HorizontalAlign="Right" Mode="NumericPages"></PagerStyle>--%>
                                    </asp:DataGrid>
                                </div>
                                <asp:Label ID="lab_Msg1" runat="server" ForeColor="Red" Visible="False" CssClass="font">查無資料!</asp:Label>
                                <%--<uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>--%>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                    <table id="DataTable2" cellspacing="1" cellpadding="1" width="100%" runat="server">
                        <tr>
                            <td class="bluecol" width="15%">功能路徑
                            </td>
                            <td class="whitecol">
                                <asp:Label ID="lFunPath" runat="server"></asp:Label>
                            </td>
                            <td class="bluecol" width="15%">功能名稱
                            </td>
                            <td class="whitecol">
                                <asp:Label ID="lFunName" runat="server"></asp:Label>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" colspan="4">
                                <div id="Div2" runat="server">
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="True" PageSize="20" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="序號">
                                                <HeaderStyle Width="5%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_SNo" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="建檔單位">
                                                <HeaderStyle Width="18%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_GroupDistID" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="群組階層">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_GroupType" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="群組名稱">
                                                <HeaderStyle Width="20%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_GroupName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="建檔者">
                                                <HeaderStyle Width="12%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_GroupCUsr" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="最後修改者">
                                                <HeaderStyle Width="12%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_GroupMUsr" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="備註">
                                                <HeaderStyle Width="18%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_GroupNote" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="啟用">
                                                <ItemStyle Width="5%" HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_Enable" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <%--<asp:TemplateColumn HeaderText="功能">
											<ItemStyle HorizontalAlign="Center" Width="18%"></ItemStyle>
											<ItemTemplate>
												<asp:LinkButton ID="btnListData1" runat="server" Text="群組查詢" CommandName="ListData1" CssClass="linkbutton"></asp:LinkButton>
											</ItemTemplate>
										</asp:TemplateColumn>--%>
                                        </Columns>
                                        <PagerStyle Font-Size="Medium" Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                    </asp:DataGrid>
                                </div>
                                <asp:Label ID="lab_Msg2" runat="server" ForeColor="Red" Visible="False" CssClass="font">查無資料!</asp:Label>
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
        <asp:HiddenField ID="Hidfunid" runat="server" />
    </form>
</body>
</html>
