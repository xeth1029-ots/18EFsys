<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_037.aspx.vb" Inherits="WDAIIP.SYS_03_037" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>學員資料整合查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function ChangeMode(num) {
            $("#MenuTable_td_1").removeClass();
            $("#MenuTable_td_2").removeClass();
            $("#MenuTable_td_3").removeClass();
            $("#MenuTable_td_4").removeClass();
            $("#tb_VIEW1").hide();
            $("#tb_VIEW2").hide();
            $("#tb_VIEW3").hide();
            $("#tb_VIEW4").hide();

            switch (num) {
                case 1:
                    $("#MenuTable_td_1").addClass("active");
                    $("#tb_VIEW1").show();
                    break;
                case 2:
                    $("#MenuTable_td_2").addClass("active");
                    $("#tb_VIEW2").show();
                    break;
                case 3:
                    $("#MenuTable_td_3").addClass("active");
                    $("#tb_VIEW3").show();
                    break;
                case 4:
                    $("#MenuTable_td_4").addClass("active");
                    $("#tb_VIEW4").show();
                    break;
                default:
                    $("#MenuTable_td_1").addClass("active");
                    $("#tb_VIEW1").show();
                    break;
            }

            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;學員資料整合查詢</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="Frametable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <div id="divSch1" runat="server">
                        <table class="table_nw" id="tb_Query" cellspacing="1" cellpadding="1" width="100%" runat="server">
                            <tr>
                                <td class="bluecol" style="width: 20%">身分證號： </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_IDNO" runat="server"></asp:TextBox></td>
                                <td class="bluecol">姓名： </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txt_NAME" runat="server"></asp:TextBox></td>
                            </tr>
                            <tr id="tr_ddl_INQUIRY_S" runat="server">
                                <td class="bluecol_need">查詢原因</td>
                                <td class="whitecol" colspan="3">
                                    <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="4" class="whitecol">
                                    <%--<asp:Label ID="labPageSize" runat="server" DESIGNTIMEDRAGDROP="30" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" Width="6%" MaxLength="2">10</asp:TextBox>--%>
                                    <asp:Button ID="btn_Query" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>&nbsp; </td>
                            </tr>
                        </table>
                        <table class="font" id="tb_List" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td align="center">
                                    <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" AllowPaging="true" CellPadding="8">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn>
                                                <HeaderStyle Width="5%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_SNo" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="身分證號">
                                                <HeaderStyle Width="11%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_IDNO" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="姓名">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_Name" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="生日">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_Birthday" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="性別">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_SexN" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn HeaderText="最後異動時間">
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="lab_LastDate" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:TemplateColumn>
                                                <HeaderStyle Width="10%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="BTNVIEW1" runat="server" Text="檢視" CommandName="BTNVIEW1" CssClass="linkbutton"></asp:LinkButton>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                    </asp:DataGrid>
                                    <asp:Label ID="lab_Msg" runat="server" ForeColor="Red">查無資料</asp:Label>
                                </td>
                            </tr>
                        </table>

                    </div>
                    <div id="divShowData1" runat="server">
                        <table class="table_nw" id="tb_DATA1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td class="whitecol" style="width: 20%"></td>
                                <td class="whitecol" style="width: 30%"></td>
                                <td class="whitecol" style="width: 20%"></td>
                                <td class="whitecol" style="width: 30%"></td>
                            </tr>
                            <tr>
                                <td class="bluecol">姓名</td>
                                <td class="whitecol">
                                    <asp:Label ID="LName" runat="server"></asp:Label></td>
                                <td class="bluecol">身分證號</td>
                                <td class="whitecol">
                                    <asp:Label ID="LIDNO" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="bluecol">生日 </td>
                                <td class="whitecol">
                                    <asp:Label ID="LBIRTH" runat="server"></asp:Label></td>
                                <td class="bluecol">性別 </td>
                                <td class="whitecol">
                                    <asp:Label ID="LSEX2" runat="server"></asp:Label></td>
                            </tr>
                            <tr>
                                <td class="whitecol" style="width: 20%"></td>
                                <td class="whitecol" style="width: 30%"></td>
                                <td class="whitecol" style="width: 20%"></td>
                                <td class="whitecol" style="width: 30%"></td>
                            </tr>
                        </table>
                        <div>
                            <table class="font" id="MenuTable" cellspacing="0" cellpadding="0" width="50%" runat="server">
                                <tr class="newlink newlink-blue">
                                    <td onclick="ChangeMode(1);" id="MenuTable_td_1">個人基本資料 </td>
                                    <td onclick="ChangeMode(2);" id="MenuTable_td_2">曾報名課程 </td>
                                    <td onclick="ChangeMode(3);" id="MenuTable_td_3">參訓學員歷史 </td>
                                    <td onclick="ChangeMode(4);" id="MenuTable_td_4">補助費用歷史 </td>
                                </tr>
                            </table>
                        </div>
                        <div>
                            <table class="table_nw" id="tb_VIEW1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td>
                                        <table class="font" id="tb_VIEW1_d1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                            <tr>
                                                <td class="whitecol" style="width: 20%"></td>
                                                <td class="whitecol" style="width: 30%"></td>
                                                <td class="whitecol" style="width: 20%"></td>
                                                <td class="whitecol" style="width: 30%"></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">身分別 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="PassPortNO" runat="server"></asp:Label></td>
                                                <td class="bluecol">身分證號碼 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="IDNO" runat="server"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">性別 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="Sex" runat="server"></asp:Label></td>
                                                <td class="bluecol">最高學歷 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="DegreeID" runat="server"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">婚姻狀況 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="MaritalStatus" runat="server"></asp:Label></td>
                                                <td class="bluecol">畢業狀況 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="GradID" runat="server"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">學校名稱 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="School" runat="server"></asp:Label></td>
                                                <td class="bluecol">科系 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="Department" runat="server"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">兵役 </td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:Label ID="MilitaryID" runat="server"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">通訊地址 </td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:Label ID="Address" runat="server"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">戶籍地址 </td>
                                                <td class="whitecol" colspan="3">
                                                    <asp:Label ID="LabHouseholdAddress" runat="server"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">聯絡電話(日) </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="Phone1" runat="server"></asp:Label></td>
                                                <td class="bluecol">聯絡電話(夜) </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="Phone2" runat="server"></asp:Label></td>
                                            </tr>
                                            <tr>
                                                <td class="bluecol">電子信箱 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="Email" runat="server"></asp:Label></td>
                                                <td class="bluecol">行動電話 </td>
                                                <td class="whitecol">
                                                    <asp:Label ID="CellPhone" runat="server"></asp:Label></td>
                                            </tr>
                                        </table>
                                        <%--<div align="left"><asp:Label ID="lab_Msg11" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label></div>
                                            <asp:DataGrid ID="Datagrid11" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8"><AlternatingItemStyle BackColor="#F5F5F5" /><HeaderStyle CssClass="head_navy" />
                                            <Columns></Columns><PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle></asp:DataGrid>--%>
                                    </td>
                                </tr>
                            </table>
                            <table class="table_nw" id="tb_VIEW2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td class="table_title" colspan="4" align="center">已報名課程</td>
                                </tr>
                                <tr>
                                    <td>
                                        <div align="left">
                                            <asp:Label ID="lab_Msg12" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                        </div>
                                        <asp:DataGrid ID="Datagrid12" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <HeaderStyle CssClass="head_navy" />
                                            <Columns>
                                                <asp:TemplateColumn HeaderText="項次" HeaderStyle-Width="5%">
                                                    <ItemStyle HorizontalAlign="Center" Width="5%"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lab_SNo12" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn DataField="YEARS" HeaderText="年度" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="PLANNAME" HeaderText="訓練計畫" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區分署" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="OCID" HeaderText="班級代碼" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="SFTDATE" HeaderText="訓練期間" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="RELENTERDATE" HeaderText="報名日期" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <%--<asp:BoundColumn DataField="ESERNUM" HeaderText="序號" HeaderStyle-Width="5%"></asp:BoundColumn>--%>
                                                <asp:BoundColumn DataField="SIGNUPSTATUS_N" HeaderText="e網審核結果" HeaderStyle-Width="8%"></asp:BoundColumn>
                                            </Columns>
                                            <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="table_title" colspan="4" align="center">取消報名課程</td>
                                </tr>
                                <tr>
                                    <td>
                                        <div align="left">
                                            <asp:Label ID="lab_Msg12b" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                        </div>
                                        <asp:DataGrid ID="Datagrid12b" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <HeaderStyle CssClass="head_navy" />
                                            <Columns>
                                                <asp:TemplateColumn HeaderText="項次" HeaderStyle-Width="5%">
                                                    <ItemStyle HorizontalAlign="Center" Width="5%"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lab_SNo12b" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn DataField="YEARS" HeaderText="年度" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="PLANNAME" HeaderText="訓練計畫" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區分署" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="OCID" HeaderText="班級代碼" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="SFTDATE" HeaderText="訓練期間" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="RELENTERDATE" HeaderText="報名日期" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <%--<asp:BoundColumn DataField="ESERNUM" HeaderText="序號" HeaderStyle-Width="5%"></asp:BoundColumn>--%>
                                                <asp:BoundColumn DataField="CANCELTIME" HeaderText="取消報名時間" HeaderStyle-Width="8%"></asp:BoundColumn>
                                            </Columns>
                                            <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                        </asp:DataGrid>
                                    </td>
                                </tr>

                            </table>
                            <table class="table_nw" id="tb_VIEW3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td>
                                        <div align="left">
                                            <asp:Label ID="lab_Msg13" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                        </div>
                                        <asp:DataGrid ID="Datagrid13" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <HeaderStyle CssClass="head_navy" />
                                            <Columns>
                                                <asp:TemplateColumn HeaderText="項次" HeaderStyle-Width="5%">
                                                    <ItemStyle HorizontalAlign="Center" Width="5%"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lab_SNo13" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn DataField="YEARS" HeaderText="年度" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="PLANNAME" HeaderText="訓練計畫" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區分署" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <%--<asp:BoundColumn DataField="OCID" HeaderText="班級代碼" HeaderStyle-Width="5%"></asp:BoundColumn>--%>
                                                <asp:BoundColumn DataField="ClassName" HeaderText="班級名稱" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="TRound" HeaderText="訓練期間" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="THours" HeaderText="受訓時數" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="WEEKS" HeaderText="上課時間" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="TFlag" HeaderText="訓練狀態" HeaderStyle-Width="8%"></asp:BoundColumn>
                                            </Columns>
                                            <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                        </asp:DataGrid>
                                    </td>
                                </tr>

                            </table>
                            <table class="table_nw" id="tb_VIEW4" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                                <tr>
                                    <td>
                                        <div align="left">
                                            <asp:Label ID="lab_Msg14" runat="server" Visible="False" ForeColor="Red">查無資料</asp:Label>
                                        </div>
                                        <asp:DataGrid ID="Datagrid14" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                            <AlternatingItemStyle BackColor="#F5F5F5" />
                                            <HeaderStyle CssClass="head_navy" />
                                            <Columns>
                                                <%--年度、轄區分署、訓練機構、班級代碼(OCID)、班級名稱、開訓日期、結訓日期、申請補助金額、預算別、審核狀態、撥款狀態、訓練狀態--%>
                                                <asp:TemplateColumn HeaderText="項次" HeaderStyle-Width="5%">
                                                    <ItemStyle HorizontalAlign="Center" Width="5%"></ItemStyle>
                                                    <ItemTemplate>
                                                        <asp:Label ID="lab_SNo14" runat="server"></asp:Label>
                                                    </ItemTemplate>
                                                </asp:TemplateColumn>
                                                <asp:BoundColumn DataField="YEARS" HeaderText="年度" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區分署" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練機構" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="OCID" HeaderText="班級代碼" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="CLASSCNAME" HeaderText="班級名稱" HeaderStyle-Width="8%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="STDATE" HeaderText="開訓日期" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="FTDATE" HeaderText="結訓日期" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="SumOfMoney" HeaderText="申請補助金額" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="BudName" HeaderText="預算別" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="AppliedStatusM" HeaderText="審核狀態" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="AppliedStatus" HeaderText="撥款狀態" HeaderStyle-Width="5%"></asp:BoundColumn>
                                                <asp:BoundColumn DataField="StudStatus" HeaderText="訓練狀態" HeaderStyle-Width="5%"></asp:BoundColumn>
                                            </Columns>
                                            <PagerStyle Font-Bold="true" HorizontalAlign="Center" Mode="NumericPages"></PagerStyle>
                                        </asp:DataGrid>
                                    </td>
                                </tr>
                                <tr>
                                    <%--撥款通過總額--%>
                                    <td class="whitecol">(補助總額：
											<asp:Label ID="LabTotal" runat="server"></asp:Label>)-(經費審核通過總額：
											<asp:Label ID="LabSumOfMoney" runat="server"></asp:Label>)=(剩餘可用額度：
											<asp:Label ID="RemainSub" runat="server"></asp:Label>)
                                    </td>
                                </tr>
                                <tr>
                                    <td class="whitecol">
                                        <asp:Label ID="LabCostDay" runat="server"></asp:Label></td>
                                </tr>
                                <%-- <tr>
                                    <td class="whitecol"><font color="red">
                                        <asp:Label ID="Lab_TipMsg2" runat="server" Text=""></asp:Label></font></td>
                                </tr>--%>
                            </table>
                        </div>
                        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <asp:Button ID="btnBACK1" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_sch1" runat="server" />
        <asp:HiddenField ID="Hid_IDNO" runat="server" />
        <asp:HiddenField ID="Hid_MSTYLE" runat="server" />
    </form>
</body>
</html>
