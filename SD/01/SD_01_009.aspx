 
<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_01_009.aspx.vb" Inherits="WDAIIP.SD_01_009" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>報名來源統計</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link type="text/css" href="../../css/style.css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript">
        function GETvalue() {
            document.getElementById('Button4').click();
        }

        function choose_class() {
            openClass('../02/SD_02_ch.aspx?RID=' + document.form1.RIDValue.value);
        }

        //function choose_other() {
        //    var OCID = document.form1.OCIDValue1.value
        //    window.open('SD_02_003_other.aspx?OCID=' + OCID, '', 'width=550,height=250,location=0,status=0,menubar=0,scrollbars=1,resizable=0');
        //}

        //function search(){
        //	if(document.form1.OCIDValue1.value==''){
        //		alert('請選擇職類班別!')
        //		return false;
        //	}
        //}
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;招生作業&gt;&gt;報名來源統計</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Panel ID="table_3" runat="server">
                        <table class="table_sch" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol" style="width: 20%">訓練機構</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                    <input id="Button8" type="button" value="..." name="Button8" runat="server" class="asp_button_Mini" />
                                    <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                                    <asp:Button ID="Button4" Style="display: none" runat="server" Text="Button4" CssClass="asp_button_S"></asp:Button>
                                    <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()"><asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">職類/班別</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                    <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                    <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                                    <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                                    <span id="HistoryList" style="position: absolute; display: none; left: 28%"><asp:Table ID="Historytable" runat="server" Width="100%"></asp:Table></span>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">通俗職類</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">開訓日期</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="STDate1" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= stdate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /> ~
                                    <asp:TextBox ID="STDate2" runat="server" MaxLength="10" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= stdate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td class="whitecol" align="center"><asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                        </table>
                    </asp:Panel>
                    <table id="table4" cellspacing="0" cellpadding="0" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" ShowFooter="true" AllowCustomPaging="true" AllowPaging="true" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="序號">
                                            <HeaderStyle Width="4%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="orgName" HeaderText="訓練單位">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="classCname" HeaderText="班級名稱">
                                            <HeaderStyle HorizontalAlign="Center" Width="19%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="chtotal" HeaderText="報名人數">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ch1" HeaderText="網路人數">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ch2" HeaderText="現場人數">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="ch3" HeaderText="通訊人數">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>
                                        <%--<asp:BoundColumn DataField="ch4" HeaderText="一般推介單人數">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="EPW" HeaderText="免試推介單人數">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>--%>
                                        <asp:BoundColumn DataField="EP2P" HeaderText="專案核定報名人數">
                                            <HeaderStyle HorizontalAlign="Center" Width="8%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <FooterStyle HorizontalAlign="Center"></FooterStyle>
                                        </asp:BoundColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr>
                            <td align="center"><uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler></td>
                        </tr>
                    </table>
                    <div align="center"><asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label></div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>