<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_03_012_add.aspx.vb" Inherits="WDAIIP.SD_03_012_add" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <%--報名暨報到名單確認--%>
    <title>參訓學員名單確認</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function savedataCHK1() {
            var rst1 = true; //正常再次檢核。
            var vMsgchk1 = "確認名單：首頁>>學員動態管理>>學員資料管理>>學員參訓，已審核錄取名單，該班不可以再進行學員參訓!!";
            rst1 = confirm(vMsgchk1);
            return rst1;
        }

        function savedataCHK2() {
            var rst1 = true; //正常再次檢核。
            var vMsgchk1 = "解鎖學員參訓作業：首頁>>學員動態管理>>學員資料管理>>學員參訓，已審核錄取名單，該班可以再進行學員參訓!!";
            rst1 = confirm(vMsgchk1);
            return rst1;
        }

        function ChangeMode(num) {
            var DataGridTable1 = document.getElementById('DataGridTable1');
            var DataGridTable2 = document.getElementById('DataGridTable2');
            //var MenuTable = document.getElementById('MenuTable');
            var BtnSave1 = document.getElementById('BtnSave1'); //儲存確認
            var BtnSave2 = document.getElementById('BtnSave2'); //解鎖
            var BtnPrint1 = document.getElementById('BtnPrint1'); //列印
            DataGridTable1.style.display = 'none';
            DataGridTable2.style.display = 'none';
            if (BtnSave1) { BtnSave1.style.display = 'none'; }
            if (BtnSave2) { BtnSave2.style.display = 'none'; }
            if (BtnPrint1) { BtnPrint1.style.display = 'none'; }
            var MenuTable_td_1 = $("#MenuTable_td_1");
            var MenuTable_td_2 = $("#MenuTable_td_2");
            MenuTable_td_1.removeClass();
            MenuTable_td_2.removeClass();
            switch (num) {
                case 1:
                    MenuTable_td_1.addClass("active");
                    DataGridTable1.style.display = '';
                    break;
                case 2:
                    MenuTable_td_2.addClass("active");
                    if (BtnSave1) { BtnSave1.style.display = ''; }
                    if (BtnSave2) { BtnSave2.style.display = ''; }
                    if (BtnPrint1) { BtnPrint1.style.display = ''; }
                    DataGridTable2.style.display = '';
                    break;
            }
            //fix 動態變動顯示內容, 會造成顯示內容超出 iframe 顯示區域被遮掉的情況 
            //if (!_isIE) { parent.setMainFrameHeight(); } //重新調整頁面高度
            //if (parent && parent.setMainFrameHeight != undefined) { parent.setMainFrameHeight(); }
            if (document.body) { window.scroll(0, document.body.scrollHeight); }
            if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <%--<table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;學員資料管理&gt;&gt;報名暨報到名單確認</asp:Label>
                                -<asp:Label ID="labActText" runat="server"></asp:Label>
                            </td>
                        </tr>
                    </table>--%>
                    <table style="cursor: pointer" id="MenuTable" class="font" border="0" width="50%" cellspacing="0" cellpadding="0" runat="server">
                        <tr class="newlink newlink-blue">
                            <%--
                            <td id="ChangeMode1a" onclick="ChangeMode(1);" background="../../images/bookmark_01.gif" width="3"></td>
						    <td id="ChangeMode1b" onclick="ChangeMode(1);" background="../../images/bookmark_02.gif" width="100" align="center">報名名單 </td>
						    <td id="ChangeMode1c" onclick="ChangeMode(1);" background="../../images/bookmark_03.gif" width="11"></td>
						    <td id="ChangeMode2a" onclick="ChangeMode(2);" background="../../images/bookmark_01.gif" width="3"></td>
						    <td id="ChangeMode2b" onclick="ChangeMode(2);" background="../../images/bookmark_02.gif" width="100" align="center"><font style="color: #009900;">報到名單</font> </td>
						    <td id="ChangeMode2c" onclick="ChangeMode(2);" background="../../images/bookmark_03.gif" width="11"></td>
                            --%>
                            <td id="MenuTable_td_1" onclick="ChangeMode(1);">報名名單</td>
                            <td id="MenuTable_td_2" onclick="ChangeMode(2);">報到名單</td>
                        </tr>
                    </table>
                    <table id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="LabClassName1" runat="server" CssClass="font"></asp:Label><br />
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <ItemStyle Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>報名序號</HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="labSeqno" runat="server"></asp:Label>
                                                <input id="ESETID" type="hidden" runat="server" />
                                                <input id="ESERNUM" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labStdName" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="身分證字號">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labIDNO" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="報名日期">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="RelENTERDATE" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="報名路徑">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="LabEnterPath" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="報名審核<br>結果">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="signUpStatusN" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="審核失敗原因">
                                            <ItemTemplate>
                                                <asp:Label ID="signUpMemo" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <asp:Label ID="LabClassName2" runat="server" CssClass="font"></asp:Label><br />
                                <asp:DataGrid ID="DataGrid2" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False">
                                    <AlternatingItemStyle BackColor="#EEEEEE" Height="20px" />
                                    <ItemStyle Height="20px" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn>
                                            <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>序號</HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:Label ID="labSeqno" runat="server"></asp:Label>
                                                <input id="SETID" type="hidden" runat="server" />
                                                <input id="EnterDate" type="hidden" runat="server" />
                                                <input id="SerNum" type="hidden" runat="server" />
                                                <input id="SOCID" type="hidden" runat="server" />
                                                <input id="HStudStatus" type="hidden" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="姓名">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labStdName" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="身分證字號">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labIDNO" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="保險證號">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labActNo" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="預算別">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labBudget" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="在保資格<br>(符合/不符合)">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labCapMode" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="離退訓">
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="labStatus23" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr id="trODNUMBER" runat="server">
                            <td class="whitecol">&nbsp;&nbsp; <font color="red">公文文號(必填)：</font><asp:TextBox ID="ODNUMBER" Width="280" runat="server" MaxLength="100"></asp:TextBox></td>
                        </tr>
                    </table>
                    <table width="100%" class="font">
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Button ID="BtnSave1" runat="server" Text="確認名單" CssClass="asp_button_M"></asp:Button>
                                &nbsp;<asp:Button ID="BtnSave2" runat="server" Text="解鎖學員參訓作業" CssClass="asp_button_M"></asp:Button>

                                &nbsp;<asp:Button ID="BtnPrint1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                &nbsp;<asp:Button ID="BtnBack1" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td align="center" class="whitecol">
                                <asp:Label ID="msg" runat="server" ForeColor="Red" CssClass="font"></asp:Label></td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="Hid_INQUIRY_SCH" runat="server" />
        <asp:HiddenField ID="Hid_CFGUID" runat="server" />
        <asp:HiddenField ID="Hid_OCID" runat="server" />
        <asp:HiddenField ID="Hid_CFSEQNO" runat="server" />        
    </form>
</body>
</html>
