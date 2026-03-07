<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_021.aspx.vb" Inherits="WDAIIP.SYS_03_021" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD html 4.0 Transitional//EN">
<html>
<head>
    <title>計畫功能設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function Show_ListSort(tmpObj, tmpText, tmpValue, tmpAct) {
            var tmpID = tmpObj.selectedIndex;
            var tmpItem = tmpObj.value;
            if (tmpAct == "up") {
                if (tmpID > 0) {
                    tmpObj.options[tmpID].text = tmpObj.options[tmpID - 1].text;
                    tmpObj.options[tmpID].value = tmpObj.options[tmpID - 1].value;
                    tmpObj.options[tmpID - 1].text = tmpText;
                    tmpObj.options[tmpID - 1].value = tmpValue;
                }
            } else {
                if (tmpID < tmpObj.length - 1) {
                    tmpObj.options[tmpID].text = tmpObj.options[tmpID + 1].text;
                    tmpObj.options[tmpID].value = tmpObj.options[tmpID + 1].value;
                    tmpObj.options[tmpID + 1].text = tmpText;
                    tmpObj.options[tmpID + 1].value = tmpValue;
                }
            }
            tmpObj.value = tmpItem;
        }

        function Show_SelectAll(tmpName1, tmpName2, tmpCnt) {
            if (document.getElementById(tmpName1)) {
                var FG_CHECK = document.getElementById(tmpName1).checked;
                //$("#DataGrid1 input[type='checkbox']").prop("checked", FG_CHECK);
                //$("#DataGrid1 input[type='checkbox']:not(#" + tmpName1 + ")").prop("checked", FG_CHECK);
                $("#DataGrid1 input[type='checkbox'][name*='chk_Enable']:not(#" + tmpName1 + ")").prop("checked", FG_CHECK);
                //for (i = 0; i < tmpCnt; i++) {
                //    if (document.getElementById(tmpName2.replace("ctl2", "ctl" + (i + 2)))) {
                //        if (document.getElementById(tmpName2.replace("ctl2", "ctl" + (i + 2))).disabled == false) {
                //            document.getElementById(tmpName2.replace("ctl2", "ctl" + (i + 2))).checked = document.getElementById(tmpName1).checked;
                //        }
                //    }
                //}
            }
        }

        function Show_SubList(itmName, tdName1, tdName2, subs, tdItems) {
            if (document.getElementById(itmName + "1").style.display == "inline" || document.getElementById(itmName + "1").style.display == "") {
                for (i = 1; i < subs; i++) {
                    document.getElementById(itmName + i).style.display = "none";
                }
                document.getElementById(itmName + "td0").rowSpan = 1;
                document.getElementById(itmName + "td1").colSpan = tdItems - 1;
                for (i = 2; i < tdItems; i++) {
                    document.getElementById(itmName + "td" + i).style.display = "none";
                }
                document.getElementById(tdName1).style.display = "none";
                document.getElementById(tdName2).style.display = "inline";
            } else {
                for (i = 1; i < subs; i++) {
                    document.getElementById(itmName + i).style.display = "inline";
                }
                document.getElementById(itmName + "td0").rowSpan = subs;
                document.getElementById(itmName + "td1").colSpan = 1;
                for (i = 2; i < tdItems; i++) {
                    document.getElementById(itmName + "td" + i).style.display = "inline";
                }
                document.getElementById(tdName1).style.display = "inline";
                document.getElementById(tdName2).style.display = "none";
            }
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;計畫功能設定</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" runat="server" id="Table1">
                        <tr>
                            <td class="bluecol" align="center" style="width: 20%">計畫功能</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddlTPlan" runat="server" AutoPostBack="True"></asp:DropDownList></td>
                        </tr>
                        <tr id="tr_Fun" runat="server">
                            <td class="bluecol" align="center">功能類別
                            </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddlFun" runat="server" AutoPostBack="True">
                                    <asp:ListItem Value="">全部</asp:ListItem>
                                    <asp:ListItem Value="TC">[TC]訓練機構管理</asp:ListItem>
                                    <asp:ListItem Value="SD">[SD]學員動態管理</asp:ListItem>
                                    <asp:ListItem Value="CP">[CP]查核績效管理</asp:ListItem>
                                    <asp:ListItem Value="TR">[TR]訓練需求管理</asp:ListItem>
                                    <asp:ListItem Value="CM">[CM]訓練經費控管</asp:ListItem>
                                    <asp:ListItem Value="FM">[FM]設備預算管理</asp:ListItem>
                                    <asp:ListItem Value="SE">[SE]技能檢定管理</asp:ListItem>
                                    <asp:ListItem Value="EXAM">[EXAM]招生甄試管理</asp:ListItem>
                                    <asp:ListItem Value="SV">[SV]問卷管理</asp:ListItem>
                                    <asp:ListItem Value="OB">[OB]委外訓練管理</asp:ListItem>
                                    <asp:ListItem Value="SYS">[SYS]系統管理</asp:ListItem>
                                    <asp:ListItem Value="FAQ">[FAQ]問答集</asp:ListItem>
                                    <asp:ListItem Value="OO">[OO]其他系統</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr id="tr_Btn2" runat="server">
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="btn_Save2" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr class="whitecol">
                            <td align="left" colspan="2">
                                <asp:DataGrid ID="DataGrid1" runat="server" AutoGenerateColumns="False" CssClass="font" Width="100%" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#EEEEEE" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="功能類別">
                                            <HeaderStyle Width="15%"></HeaderStyle>
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:TextBox ID="txtFunID" Visible="False" runat="server"></asp:TextBox>
                                                <asp:Label ID="lab_MainMenu" runat="server" ForeColor="#00007F"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能項目">
                                            <HeaderStyle Width="50%"></HeaderStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_FunName" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="Memo" HeaderText="備註">
                                            <HeaderStyle Width="25%"></HeaderStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="選用">
                                            <HeaderStyle Width="10%" />
                                            <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            <HeaderTemplate>
                                                <asp:CheckBox ID="chk_EnableAll" runat="server" Text="選用"></asp:CheckBox>
                                            </HeaderTemplate>
                                            <ItemTemplate>
                                                <asp:CheckBox ID="chk_Enable" runat="server" Text="選用"></asp:CheckBox>&nbsp;
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                        <tr id="tr_Btn" runat="server">
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="btn_Save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>


                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
