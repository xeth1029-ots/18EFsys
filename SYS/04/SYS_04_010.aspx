<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SYS_04_010.aspx.vb" Inherits="WDAIIP.SYS_04_010" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>核銷關帳設定</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        //判斷儲存
        function chkSave() {
            var msg = '';
            var txtClose1 = document.getElementById('txtClose1');
            var ddlClose2M = document.getElementById('ddlClose2M');
            var ddlClose2D = document.getElementById('ddlClose2D');
            var txtClose3 = document.getElementById('txtClose3');
            var txtClose4 = document.getElementById('txtClose4');
            var txtClose5 = document.getElementById('txtClose5');
            var ddlClose6M = document.getElementById('ddlClose6M');
            var ddlClose6D = document.getElementById('ddlClose6D');

            if (txtClose1.value == '') msg += '請輸入自辦跨年度月數!\n';
            else if (!isUnsignedInt(txtClose1.value)) msg += '自辦跨年度月數輸入格式有誤!\n';

            if (ddlClose2M.value == '') msg += '請選擇自辦當年度月份!\n';
            if (ddlClose2D.value == '') msg += '請選擇自辦當年度日期!\n';

            if (ddlClose2M.value != '' && ddlClose2D.value != '') {
                switch (ddlClose2M.value) {
                    //2月 
                    case '2':
                        if (parseInt(ddlClose2D.value) > 29) msg += '自辦當年度月份/日期不符合!\n';
                        break;

                    //大月 
                    case '1':
                    case '3':
                    case '5':
                    case '7':
                    case '8':
                    case '10':
                    case '12':

                        break;

                    //小月 
                    default:
                        if (parseInt(ddlClose2D.value) > 30) msg += '自辦當年度月份/日期不符合!\n';
                        break;
                }
            }

            if (txtClose3.value == '') msg += '請輸入委辦跨年度委訓部份月數!\n';
            else if (!isUnsignedInt(txtClose3.value)) msg += '委辦跨年度委訓部份月數輸入格式有誤!\n';

            //if (txtClose4.value == '') msg += '請輸入委辦跨年度中心部份月數!\n';
            if (txtClose4.value == '') msg += '請輸入委辦跨年度分署部份月數!\n';
            //else if (!isUnsignedInt(txtClose3.value)) msg += '委辦跨年度中心部份月數輸入格式有誤!\n';
            else if (!isUnsignedInt(txtClose3.value)) msg += '委辦跨年度分署部份月數輸入格式有誤!\n';

            if (txtClose5.value == '') msg += '請輸入委辦當年度委訓部份月數!\n';
            else if (!isUnsignedInt(txtClose3.value)) msg += '委辦當年度委訓部份月數輸入格式有誤!\n';

            if (ddlClose6M.value == '') msg += '請選擇委辦當年度委訓部份月份!\n';
            if (ddlClose6D.value == '') msg += '請選擇委辦當年度委訓部份日期!\n';

            if (ddlClose6M.value != '' && ddlClose6D.value != '') {
                switch (ddlClose6M.value) {
                    //2月 
                    case '2':
                        if (parseInt(ddlClose6D.value) > 29) msg += '委辦當年度委訓部份月份/日期不符合!\n';
                        break;

                    //大月 
                    case '1':
                    case '3':
                    case '5':
                    case '7':
                    case '8':
                    case '10':
                    case '12':

                        break;

                    //小月 
                    default:
                        if (parseInt(ddlClose6D.value) > 30) msg += '委辦當年度委訓部份月份/日期不符合!\n';
                        break;
                }
            }

            if (msg != '') {
                alert(msg);
                return false;
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
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;核銷關帳設定</asp:Label>
                </td>
            </tr>
        </table>
    <table id="FrameTable1" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td>
               <%-- <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;系統管理&gt;&gt;系統參數管理&gt;&gt;<font color="#990000">核銷關帳設定</font>
                        </td>
                    </tr>
                </table>--%>
                <table class="table_sch" cellspacing="1" cellpadding="1" width="100%" >
                    <tr>
                        <td class="bluecol" style="width:20%">
                            訓練計畫
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="ddlTPlan" runat="server" AutoPostBack="True">
                            </asp:DropDownList>
                        </td>
                    </tr>
                    <tr id="trSys" runat="server">
                        <td class="bluecol">
                            系統預設
                        </td>
                        <td class="whitecol">
                            <asp:CheckBox ID="chkSys" AutoPostBack="True" runat="server"></asp:CheckBox>
                        </td>
                    </tr>
                </table>
                <br>
                <input type="hidden" id="hidSBCID" runat="server">
                <table class="table_sch" runat="server" id="tbList" cellpadding="1" cellspacing="1">
                    <tr>
                        <td class="bluecol" style="width:20%">
                            訓練分類
                        </td>
                        <td class="bluecol" style="width:10%">
                            執行年度
                        </td>
                        <td class="bluecol" style="width:70%">
                            關帳時間
                        </td>
                    </tr>
                    <tr>
                        <td class="whitecol" rowspan="2">
                            自辦<font style="color: red">*</font>
                        </td>
                        <td class="whitecol" align="center">
                            跨年度
                        </td>
                        <td class="whitecol">
                            &nbsp;依班級結訓日後依班級結訓日後
                            <asp:TextBox ID="txtClose1" runat="server" Columns="1" Width="8%"></asp:TextBox>個月
                        </td>
                    </tr>
                    <tr>
                        <td class="whitecol" align="center">
                            當年度
                        </td>
                        <td class="whitecol">
                            &nbsp;隔年
                            <asp:DropDownList ID="ddlClose2M" runat="server">
                            </asp:DropDownList>
                            月
                            <asp:DropDownList ID="ddlClose2D" runat="server">
                            </asp:DropDownList>
                            日
                        </td>
                    </tr>
                    <tr>
                        <td class="whitecol" rowspan="2">
                            &nbsp;&nbsp;委辦(包含補助、合辦)<font style="color: red">*</font>
                        </td>
                        <td class="whitecol" align="center">
                            跨年度
                        </td>
                        <td class="whitecol">
                            &nbsp;委訓部份:依班級結訓日後
                            <asp:TextBox ID="txtClose3" runat="server" Columns="1" Width="8%"></asp:TextBox>個月<br>
                            <%--&nbsp;中心部份:依班級結訓日後--%>
                            &nbsp;分署部份:依班級結訓日後
                            <asp:TextBox ID="txtClose4" runat="server" Columns="1" Width="8%"></asp:TextBox>個月
                        </td>
                    </tr>
                    <tr>
                        <td class="whitecol" align="center">
                            當年度
                        </td>
                        <td class="whitecol">
                            &nbsp;委訓部份:依班級結訓日後
                            <asp:TextBox ID="txtClose5" runat="server" Columns="1" Width="8%"></asp:TextBox>個月<br>
                            <%--&nbsp;中心部份:隔年--%>
                            &nbsp;分署部份:隔年
                            <asp:DropDownList ID="ddlClose6M" runat="server">
                            </asp:DropDownList>
                            月
                            <asp:DropDownList ID="ddlClose6D" runat="server">
                            </asp:DropDownList>
                            日
                        </td>
                    </tr> 
                </table>
                <table width="100%">
                    <tr>
                        <td align="center" colspan="3" class="whitecol">                            
                            <asp:Button ID="btnSave" runat="server" Text="儲存" CssClass="asp_button_M"></asp:Button>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
