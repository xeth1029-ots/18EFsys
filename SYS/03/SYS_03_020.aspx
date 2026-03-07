<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_03_020.aspx.vb" Inherits="WDAIIP.SYS_03_020" %>

<!DOCTYPE html PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>功能頁面設定</title>
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
        function Show_ListSort(tmpObj, tmpAct) {
            var tmpID = tmpObj.selectedIndex;
            if (tmpID != "-1") {
                var tmpText = tmpObj.options[tmpID].text;
                var tmpValue = tmpObj.options[tmpID].value;
                if (tmpAct == "up") {
                    if (tmpID > 0) {
                        tmpObj.add(new Option(tmpText, tmpValue), tmpID - 1);
                        tmpObj.remove(tmpID + 1);
                    }
                } else {
                    if (tmpID < tmpObj.length - 1) {
                        tmpObj.add(new Option(tmpText, tmpValue), tmpID + 2);
                        tmpObj.remove(tmpID);
                    }
                }
                tmpObj.value = tmpValue;
                document.getElementById("hide_Sort").value = tmpObj.selectedIndex;
            }
            return false;
        }

        function Check_Data() {
            var errMsg = "";
            if (document.getElementById("list_MainMenu").selectedIndex == 0) {
                errMsg += "請選擇功能類型。\n";
            }
            if (document.getElementById("txt_FunName").value == "") {
                errMsg += "請輸入功能名稱。\n";
            }
            if (errMsg == "") {
                return true;
            } else {
                alert(errMsg);
                return false;
            }
        }

        function TPlanIDFunSelected(obj, sptshow) {
            //debugger; '';//'inline';
            document.getElementById(sptshow).style.display = obj.checked ? '' : 'none';
        }

        //選擇全部
        function SelectAll2(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }

        //選擇全部
        function SelectAll(obj, hidobj, othobj) {
            var num = getCheckBoxListValue(obj).length;
            var myallcheck = document.getElementById(obj + '_' + 0);
            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    //mycheck.checked=myallcheck.checked;
                    //debugger;
                    //下列為本程特別規定 Start;		
                    var result = document.getElementById(othobj).value;
                    if (!myallcheck.checked) {
                        if (result.indexOf(',' + mycheck.id + ',') != -1) {
                            mycheck.checked = true;
                        }
                        else {
                            mycheck.checked = myallcheck.checked;
                        }
                    }
                    else {
                        mycheck.checked = myallcheck.checked;
                    }
                    //上列為本程特別規定 End;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;功能權限管理&gt;&gt;功能頁面設定</asp:Label>
                </td>
            </tr>
        </table>
        <table class="font" id="FrameTable2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="tb_View" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td>
                                <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                                    <tr>
                                        <td class="bluecol" align="center" style="width: 20%">功能類型</td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="list_MainMenu2" runat="server" AutoPostBack="true">
                                                <asp:ListItem Value="全部">全部</asp:ListItem>
                                            </asp:DropDownList>
                                            &nbsp;<asp:Button ID="btn_Add" runat="server" Text="新增功能" CssClass="asp_button_M"></asp:Button>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" align="center" width="10%">功能項目類別</td>
                                        <td class="whitecol">
                                            <asp:DropDownList ID="list_MainMenu3" runat="server" AutoPostBack="true"></asp:DropDownList></td>
                                    </tr>
                                   <%-- <tr>
                                        <td class="bluecol">功能位置</td>
                                        <td class="whitecol">
                                            <asp:TextBox ID="sch_txSPAGE" runat="server" MaxLength="100" Columns="50"></asp:TextBox>
                                            &nbsp;<asp:Button ID="btnSearch1" runat="server" Text="查詢" />
                                        </td>
                                    </tr>--%>
                                    <tr>
                                        <td class="bluecol_need">刪除聯結</td>
                                        <td class="whitecol" colspan="3">
                                            <asp:CheckBox ID="CHK_DEL_REL_2" runat="server" Text="[功能]若點選刪除，將刪除相關連結"></asp:CheckBox>
                                        </td>
                                    </tr>

                                </table>
                                <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                                    <AlternatingItemStyle BackColor="#F5F5F5" />
                                    <HeaderStyle CssClass="head_navy" />
                                    <Columns>
                                        <asp:TemplateColumn HeaderText="功能類別">
                                            <ItemStyle Width="10%"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_MainMenu" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能項目">
                                            <ItemStyle Width="29%"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_FunName" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:BoundColumn DataField="Memo" HeaderText="備註">
                                            <ItemStyle Width="22%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="SPage" HeaderText="路徑">
                                            <ItemStyle Width="20%"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:BoundColumn DataField="Sort" HeaderText="排序">
                                            <ItemStyle Width="4%" HorizontalAlign="Center"></ItemStyle>
                                        </asp:BoundColumn>
                                        <asp:TemplateColumn HeaderText="啟用">
                                            <ItemStyle Width="4%" HorizontalAlign="Center"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:Label ID="lab_Valid" runat="server"></asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                        <asp:TemplateColumn HeaderText="功能">
                                            <ItemStyle Width="11%" HorizontalAlign="Center" Font-Size="Small"></ItemStyle>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btn_Edit" CommandName="EDIT1" CssClass="linkbutton" runat="server">修改</asp:LinkButton>
                                                <asp:LinkButton ID="btn_Del" CommandName="DEL1" CssClass="linkbutton" runat="server">刪除</asp:LinkButton>
                                                <asp:LinkButton ID="BtnPrint1" CommandName="PRINT1" CssClass="linkbutton" runat="server">報表</asp:LinkButton>
                                                <input id="hide_Subs2" type="hidden" name="FunID" runat="server" />
                                            </ItemTemplate>
                                        </asp:TemplateColumn>
                                    </Columns>
                                </asp:DataGrid>
                            </td>
                        </tr>
                    </table>
                    <table id="tb_Edit" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                        <tr>
                            <td class="whitecol" align="center" colspan="4">
                                <asp:Button ID="btn_Save" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>&nbsp;
							    <asp:Button ID="btn_LoadBack" runat="server" Text="回上一頁" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" style="width: 20%">功能類型
                            </td>
                            <td class="whitecol" style="width: 30%">
                                <asp:DropDownList ID="list_MainMenu" runat="server" AutoPostBack="true">
                                    <asp:ListItem Value="==請選擇==">==請選擇==</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                            <td class="bluecol_need" style="width: 20%">功能名稱</td>
                            <td class="whitecol" style="width: 30%">
                                <asp:TextBox ID="txt_FunName" runat="server" MaxLength="25"></asp:TextBox>
                                <input id="hide_FunID" type="hidden" name="FunID" runat="server" />
                                <input id="hide_Subs" type="hidden" name="FunID" runat="server" />
                                <input id="hide_Kind" type="hidden" name="FunID" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need" style="width: 20%">功能路徑</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txt_FunRoot" runat="server" MaxLength="50" Columns="60" Width="70%"></asp:TextBox>
                                <asp:Button Style="z-index: 0" ID="btnTest1" runat="server" Text="測試" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">父功能</td>
                            <td class="whitecol">
                                <asp:DropDownList ID="list_ParentFun" runat="server" AutoPostBack="true"></asp:DropDownList></td>
                            <td class="bluecol_need">啟用狀態</td>
                            <td class="whitecol">
                                <asp:CheckBox ID="CHK_VALID" runat="server" Checked="true" Text="若為停用，請取消勾選"></asp:CheckBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">排序</td>
                            <td class="whitecol" colspan="3">
                                <table style="height: 100%" cellspacing="0" cellpadding="0" border="0">
                                    <tr>
                                        <td rowspan="2">
                                            <asp:ListBox ID="list_Sort" runat="server" Rows="8" Height="150"></asp:ListBox></td>
                                        <td valign="top">
                                            <input id="btn_Up" type="button" value="▲" name="btn_Up" runat="server" /></td>
                                    </tr>
                                    <tr>
                                        <td valign="bottom">
                                            <input id="btn_Down" type="button" value="▼" name="btn_Down" runat="server" /></td>
                                    </tr>
                                </table>
                                <input id="hide_Sort" type="hidden" name="hide_Sort" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">刪除聯結</td>
                            <td class="whitecol" colspan="3">
                                <asp:CheckBox ID="CHK_DEL_REL_1" runat="server" Text="若為停用，儲存時-將刪除相關連結"></asp:CheckBox>
                            </td>
                        </tr>

                        <tr>
                            <td class="bluecol">備註</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="txt_Note" runat="server" MaxLength="50" Columns="60"></asp:TextBox></td>
                        </tr>
                        <tr id="tr_fun" runat="server">
                            <td class="bluecol">功能</td>
                            <td class="whitecol" colspan="3">
                                <asp:CheckBoxList ID="chk_Option" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="新增" Selected="true">新增</asp:ListItem>
                                    <asp:ListItem Value="刪除" Selected="true">刪除</asp:ListItem>
                                    <asp:ListItem Value="修改" Selected="true">修改</asp:ListItem>
                                    <asp:ListItem Value="查詢" Selected="true">查詢</asp:ListItem>
                                    <asp:ListItem Value="列印" Selected="true">列印</asp:ListItem>
                                    <asp:ListItem Value="常用表件" Selected="true">常用表件</asp:ListItem>
                                </asp:CheckBoxList>
                            </td>
                            <%--
						    <td style="COLOR: white" bgcolor="#96b5e3">&nbsp;&nbsp;&nbsp; 所有計畫</td>
						    <td bgcolor="#ebf3fe">
                                <asp:checkbox id="chk_AutoAdd" runat="server" Text="加入"></asp:checkbox>
								<asp:checkbox id="chk_RemoveAll" runat="server" Text="移除" Visible="False"></asp:checkbox>
                            </td>
                            --%>
                        </tr>
                        <tr>
                            <td class="bluecol">功能2</td>
                            <td class="whitecol" colspan="3">&nbsp;<asp:CheckBox ID="TPlanIDFun" runat="server" Text="儲存時將此功能 放置 所有計畫"></asp:CheckBox><br>
                                排除下列計畫：<br />
                                <span class="SYS_td1" id="SPTPlanID2" runat="server">
                                    <asp:CheckBoxList ID="cblTPlanID2" runat="server" CssClass="font" RepeatColumns="5" RepeatDirection="Horizontal"></asp:CheckBoxList></span>
                                <input id="HidTPlanID2" type="hidden" value="0" name="HidTPlanID2" runat="server" />
                                <input id="OthTPlanID2" type="hidden" value="0" name="OthTPlanID2" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">功能3</td>
                            <td class="whitecol" colspan="3">&nbsp;<asp:CheckBox ID="TPlanIDFunDEL" runat="server" Text="儲存時將此功能 移除 所有計畫"></asp:CheckBox><br>
                                加入 下列計畫：<br />
                                <span class="SYS_td1" id="SPTPlanID3" runat="server">
                                    <asp:CheckBoxList ID="cblTPlanID3" runat="server" CssClass="font" RepeatColumns="5" RepeatDirection="Horizontal"></asp:CheckBoxList></span>
                                <input id="HidTPlanID3" type="hidden" value="0" name="HidTPlanID3" runat="server" />
                                <input id="OthTPlanID3" type="hidden" value="0" name="OthTPlanID3" runat="server" />
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">功能計畫(顯示用)</td>
                            <td class="whitecol" colspan="3">&nbsp;&nbsp;<br />
                                已加入 下列計畫：<br />
                                <span class="SYS_td1" id="SPTPlanIDhave" runat="server">
                                    <asp:CheckBoxList ID="cblTPlanIDhave" runat="server" CssClass="font" RepeatColumns="5" RepeatDirection="Horizontal" Height="30px"></asp:CheckBoxList></span>
                                <asp:Button Style="z-index: 0" ID="Btn_use1" runat="server" Text="使用功能移除" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">功能群組</td>
                            <td class="whitecol" colspan="3">&nbsp;&nbsp;<br />
                                已加入 下列群組：<br />
                                <span class="SYS_td1" id="SpAuthGroup" runat="server">
                                    <asp:CheckBoxList ID="cblAuthGroup" runat="server" CssClass="font" RepeatColumns="5" RepeatDirection="Horizontal"></asp:CheckBoxList></span>
                                <input id="HidAuthGroup" type="hidden" value="0" name="HidAuthGroup" runat="server" />
                                <asp:Button Style="z-index: 0" ID="Btn_use3" runat="server" Text="測試群組" CssClass="asp_button_M"></asp:Button>
                                <asp:Button Style="z-index: 0" ID="Btn_use2" runat="server" Text="使用群組增減" CssClass="asp_button_M"></asp:Button>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
