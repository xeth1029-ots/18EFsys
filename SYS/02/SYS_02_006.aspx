<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_02_006.aspx.vb" Inherits="WDAIIP.SYS_02_006" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SYS_02_006</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script language="javascript" type="text/javascript">
        function search() {
            if (document.form1.Account.selectedIndex == 0) {
                alert('請選擇帳號');
                return false;
            }
        }

        //        str = "2,2,3,5,6,6"; //这是一字符串
        //        var strs = new Array(); //定义一数组
        //        strs = str.split(","); //字符分割      
        //        for (i = 0; i < strs.length; i++) {
        //            document.write(strs[i] + "<br/>");    //分割后的字符输出
        //        }
        function SelectRtn(flag, objName) {
            var mycheck = document.getElementById(objName);
            if (!flag) {
                if (mycheck) { mycheck.checked = false; }
            }
        }

        function SelectAll(flag, num) {
            var hidMYChkBoxValue = document.getElementById('hidMYChkBoxValue');
            var MyChkBoxStrs = new Array();
            MyChkBoxStrs = hidMYChkBoxValue.value.split(",");
            for (i = 0; i < MyChkBoxStrs.length; i++) {
                var mycheck = document.getElementById(MyChkBoxStrs[i])
                if (mycheck) {
                    mycheck.checked = flag;
                }
            }


            //            for (var i = 0; i < num; i++) {
            //                var mycheck = document.getElementById('DataList1__ctl' + i + '_ClassName')
            //                if (mycheck) {
            //                    mycheck.checked = flag; 
            //                }
            //            }
        }
    </script>
    <style type="text/css">
        .auto-style1 { color: Black; text-align: right; padding: 4px 6px; background-color: #f1f9fc; border-right: 3px solid #49cbef; height: 36px; }
        .auto-style2 { color: #333333; padding: 4px; height: 36px; }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;帳號班級賦予</asp:Label>
                </td>
            </tr>
        </table>
        <font face="新細明體">
            <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td align="center" width="100%">
                        <%--<table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                首頁&gt;&gt;系統管理&gt;&gt;帳號維護管理&gt;&gt;<font color="#990000">帳號班級賦予</font>
                            </td>
                        </tr>
                    </table>--%>
                        <table id="Table3" class="table_sch" cellpadding="1" cellspacing="1" width="100%">
                            <tr>
                                <td class="bluecol" width="20%">年度</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="DDLYears" runat="server" AutoPostBack="True">
                                    </asp:DropDownList>

                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">計畫
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="DDLPlan" runat="server" AutoPostBack="True">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol" width="20%">帳號
                                </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="Account" runat="server">
                                    </asp:DropDownList>

                                </td>
                            </tr>

                            <tr>
                                <td class="bluecol" width="20%">核准狀態
                                </td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="Situation" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                        <asp:ListItem Value="已核准" Selected="True">已核准</asp:ListItem>
                                        <asp:ListItem Value="未核准">未核准</asp:ListItem>
                                        <asp:ListItem Value="全部">全部</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                        <table class="font" id="Table5" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <input id="Checkbox1" type="checkbox" name="Checkbox1" runat="server">全選
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataList ID="DataList1" runat="server" RepeatDirection="Horizontal" ShowFooter="False" ShowHeader="False" GridLines="Both" BorderWidth="1px" CellPadding="8" CssClass="font" Width="100%" RepeatColumns="3">
                                        <ItemTemplate>
                                            <table id="Table4" cellspacing="1" cellpadding="1" width="100%" border="0">
                                                <tr>
                                                    <td id="O">
                                                        <asp:CheckBox ID="ClassName" runat="server" CssClass="font"></asp:CheckBox><input id="OCID" type="hidden" runat="server">
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td align="right">
                                                        <asp:Label ID="Teacher" runat="server" CssClass="font"></asp:Label>
                                                    </td>
                                                </tr>
                                            </table>
                                        </ItemTemplate>
                                    </asp:DataList>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="Button2" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                        <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                    </td>
                </tr>
            </table>
        </font>
        <input type="hidden" id="hidMYChkBoxValue" runat="server" />
    </form>
</body>
</html>
