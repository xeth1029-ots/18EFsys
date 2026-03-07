<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_09_018_R.aspx.vb" Inherits="WDAIIP.SD_09_018_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>職災保加退保申報表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function SetOneOCID() { document.getElementById('Button7').click(); }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            if (!RIDValue) { return; }
            var RID = RIDValue.value;
            openClass('../../SD/02/SD_02_ch.aspx?RID=' + RID);
        }

        //CheckboxAll
        function ChangeAll(obj) {
            var objLen = document.form1.length;
            for (var iCount = 0; iCount < objLen; iCount++) {
                if (document.form1.elements[iCount].type == "checkbox") {
                    var mycheck = document.form1.elements[iCount];
                    if (!mycheck.disabled) {
                        mycheck.checked = (obj.checked == true ? true : false);
                    }
                }
            }
        }

        function GETvalue() { document.getElementById('Button7').click(); }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;職災保加退保申報表</asp:Label>
                </td>
            </tr>
        </table>
        <table id="tb_CLASSSHOW1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <table class="table_sch" id="Table3" cellspacing="1" cellpadding="1">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="center" runat="server" Width="60%" onfocus="this.blur()"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                <input id="Button5" type="button" value="..." name="Button5" runat="server" class="button_b_Mini" />
                                <asp:Button ID="Button7" Style="display: none" runat="server" Text="Button7"></asp:Button>
                                <span id="HistoryList2" style="z-index: 100; position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">職類/班別
                            </td>
                            <td class="whitecol">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini" />
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                <span id="HistoryList" style="z-index: 101; position: absolute; display: none; left: 270px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <%--<tr>
                            <td class="bluecol">期別</td>
                            <td class="whitecol"><asp:TextBox ID="CyclType" runat="server"   Width="15%" MaxLength="2"></asp:TextBox></td>
                        </tr>--%>
                        <tr>
                            <td class="bluecol">列印狀態</td>
                            <td class="whitecol">
                                <asp:RadioButtonList ID="rblTYPE1" RepeatColumns="6" CssClass="font" runat="server">
                                    <asp:ListItem Value="1" Selected="True">正面</asp:ListItem>
                                    <asp:ListItem Value="2">背面</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>

                    </table>
                    <table width="100%">
                        <tr>
                            <td class="whitecol">
                                <p align="center">
                                    <asp:Label ID="labPageSize" runat="server" ForeColor="SlateBlue">顯示列數</asp:Label>
                                    <asp:TextBox ID="TxtPageSize" runat="server" MaxLength="2" Width="6%">10</asp:TextBox>
                                    <asp:Button ID="BtnQuery" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol">
                                <p align="center">
                                    <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                                </p>
                            </td>
                        </tr>
                    </table>
                    <table id="tb_DataGrid1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AllowPaging="True" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <Columns>
                                        <asp:BoundColumn DataField="YEARS" HeaderText="年度">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="DISTNAME" HeaderText="轄區">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="PLANNAME" HeaderText="計畫">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ORGNAME" HeaderText="訓練單位">
                                            <HeaderStyle HorizontalAlign="Center" Width="10%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="STDATE" HeaderText="訓練起日">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="FTDATE" HeaderText="訓練迄日">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="11%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lkBtnPrint1" runat="server" Text="加保列印" CommandName="PRINT1" CssClass="linkbutton"></asp:LinkButton>
                                                <asp:LinkButton ID="lkBtnPrint2" runat="server" Text="退保列印" CommandName="PRINT2" CssClass="linkbutton"></asp:LinkButton>
                                                <asp:LinkButton ID="lkBtnExp1" runat="server" Text="加保匯出" CommandName="EXPORT1" CssClass="asp_Export_M"></asp:LinkButton>
                                                <asp:LinkButton ID="lkBtnExp2" runat="server" Text="退保匯出" CommandName="EXPORT2" CssClass="asp_Export_M"></asp:LinkButton>
                                                <%--<asp:LinkButton ID="lkBtnSELSTD1" runat="server" Text="挑選學員列印" CommandName="SELSTD1" CssClass="linkbutton"></asp:LinkButton>--%>
                                                <asp:LinkButton ID="lkBtnREMARKS1" runat="server" Text="【備註】設定" CommandName="REMARKS1" CssClass="linkbutton"></asp:LinkButton>
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
                        <tr>
                            <td>
                                <p align="center">&nbsp;</p>
                            </td>
                        </tr>
                    </table>


                </td>
            </tr>
        </table>

        <table id="tb_SELSTD1_1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <p align="center">
                        <asp:Label ID="labMsg2" runat="server" ForeColor="Red"></asp:Label>
                    </p>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <table id="tb_DataGrid2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td align="left" class="whitecol">
                                <div><asp:Label ID="labblue1" runat="server" ForeColor="Blue"> 姓名藍色表該學員尚有自辦在職或接受委託訓練課程在訓中</asp:Label></div>
                                <div><asp:Label ID="labgreen1" runat="server" ForeColor="Green"> * 該學員於此班已離退訓</asp:Label></div>
                                <div><asp:Label ID="LabMsgSHOW2" runat="server"></asp:Label></div>
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="7%"></HeaderStyle>
                                            <HeaderTemplate>選取<input id="CheckboxAll" type="checkbox" runat="server" /></HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <input id="Checkbox1" type="checkbox" runat="server" />
                                                <asp:HiddenField ID="Hid_SOCID" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="STUDENTID" HeaderText="學號">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="Name" HeaderText="姓名">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>--%>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <HeaderTemplate>姓名</HeaderTemplate>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate><asp:Label ID="labgreen2" runat="server" ForeColor="Green"> * </asp:Label><asp:Label ID="labName" runat="server">labName</asp:Label></ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="lkBtnPrint3" runat="server" Text="加保列印" CommandName="PRINT3" CssClass="linkbutton"></asp:LinkButton>
                                                <asp:LinkButton ID="lkBtnPrint4" runat="server" Text="退保列印" CommandName="PRINT4" CssClass="linkbutton"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="center">
                                <p align="center">
                                    <asp:Button ID="BtnPrint5" runat="server" Text="群組加保列印" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="BtnPrint6" runat="server" Text="群組退保列印" CssClass="asp_button_M"></asp:Button>

                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td class="whitecol" align="center">
                    <p align="center">
                        <asp:Button ID="BtnBACK2" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                    </p>
                </td>
            </tr>
            <%-- <tr><td align="center"><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler></td></tr> --%>
        </table>

        <table id="tb_EDITDATA3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
            <tr>
                <td>
                    <p align="center">&nbsp;</p>
                    <%--<p align="center"><asp:Label ID="labMsg3" runat="server" ForeColor="Red"></asp:Label></p>--%>
                </td>
            </tr>
            <tr>
                <td align="center">
                    <table class="table_nw" id="tb_DATA3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="whitecol" style="width: 20%"></td>
                            <td class="whitecol" style="width: 30%"></td>
                            <td class="whitecol" style="width: 20%"></td>
                            <td class="whitecol" style="width: 30%"></td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="left" colspan="4">
                                <p align="left">
                                    <asp:Label ID="labMsgORGNAME3" runat="server"></asp:Label>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="left" colspan="4">
                                <p align="left">
                                    <asp:Label ID="labMsgCLASSCNAME3" runat="server"></asp:Label>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td class="whitecol" align="left" colspan="4">
                                <p align="left">
                                    <asp:Label ID="labMsg3b" runat="server">(各欄上限500個字)</asp:Label>
                                </p>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">備註欄設定(加保)</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="NOTEINSUR" runat="server" Width="95%" MaxLength="500"></asp:TextBox>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">備註欄設定(退保)</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="NOTESURR" runat="server" Width="95%" MaxLength="500"></asp:TextBox>
                            </td>
                        </tr>

                        <tr>
                            <td class="whitecol" align="center" colspan="4">
                                <p align="center">
                                    <asp:Button ID="BTNSAVE3" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                    <asp:Button ID="BtnBACK3" runat="server" Text="回上頁" CssClass="asp_button_M"></asp:Button>
                                </p>
                            </td>
                        </tr>
                    </table>

                </td>
            </tr>
            <tr>
                <td>
                    <p align="center">&nbsp;</p>
                </td>
            </tr>
            <%-- <tr><td align="center"><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler></td></tr> --%>
        </table>

        <asp:HiddenField ID="Hid_rblTYPE1" runat="server" />
        <asp:HiddenField ID="Hid_OCID1" runat="server" />
        <asp:HiddenField ID="Hid_MSD" runat="server" />
        <asp:HiddenField ID="Hid_INSUR" runat="server" />
    </form>
</body>
</html>
