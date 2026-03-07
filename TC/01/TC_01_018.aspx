<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_01_018.aspx.vb" Inherits="WDAIIP.TC_01_018" %>

 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>材料品項資料</title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript">
        //選擇全部
        function SelectAll(obj, hidobj) {
            var num = getCheckBoxListValue(obj).length; //長度
            var myallcheck = document.getElementById(obj + '_' + 0); //第1個

            if (document.getElementById(hidobj).value != getCheckBoxListValue(obj).charAt(0)) {
                document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(0);
                for (var i = 1; i < num; i++) {
                    var mycheck = document.getElementById(obj + '_' + i);
                    mycheck.checked = myallcheck.checked;
                }
            }
            else {
                for (var i = 1; i < num; i++) {
                    if ('0' == getCheckBoxListValue(obj).charAt(i)) {
                        document.getElementById(hidobj).value = getCheckBoxListValue(obj).charAt(i);
                        var mycheck = document.getElementById(obj + '_' + i);
                        myallcheck.checked = mycheck.checked;
                        break;
                    }
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
    <table cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td>
                <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                            <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;訓練機構管理&gt;&gt;基本資料設定&gt;&gt;<font color="#990000">材料品項資料</font></asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table class="table_sch" id="Table1_sch" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td class="bluecol_need" width="100">
                            年度
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="yearlist" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            轄區
                        </td>
                        <td class="whitecol">
                            <asp:CheckBoxList ID="DistrictList" runat="server" CssClass="font" RepeatDirection="Horizontal" RepeatColumns="5">
                            </asp:CheckBoxList>
                            <input id="DistHidden" type="hidden" value="0" name="DistHidden" runat="server">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            品名
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="CNAME" runat="server" Width="210px"></asp:TextBox>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <table class="font" id="Table6_exp" cellspacing="0" cellpadding="0" width="100%" border="0">
                    <tr align="center">
                        <td>
                            <asp:Button ID="bt_export" runat="server" Text="匯出" CssClass="asp_Export_M"></asp:Button>
                        </td>
                    </tr>
                </table>
                <p align="center">
                    <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label></p>
            </td>
        </tr>
        <tr>
            <td>
                <table id="ResultTable" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            <div id="Div1" runat="server">
                                <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AllowPaging="True">
                                    <AlternatingItemStyle BackColor="#f5f5f5"></AlternatingItemStyle>
                                    <HeaderStyle CssClass="head_navy"></HeaderStyle>
                                    <PagerStyle Visible="False"></PagerStyle>
                                </asp:DataGrid></div>
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
    </form>
</body>
</html>
