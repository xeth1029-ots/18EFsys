<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Calendar.aspx.vb" Inherits="WDAIIP.Calendar" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>請選擇日期</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <script src="../js/common.js" type="text/javascript"></script>
    <script type="text/javascript">
        //取得開結訓日期
        var STDate = new Date(getParamValue('STDate'));
        var FTDate = new Date(getParamValue('FTDate'));
        var NowDate = new Date(getParamValue('NowDate'));
        var btn = getParamValue('Button');
        var ValueField = getParamValue('ValueField');

        function Show_Calendar() {
            var oNowMonth = document.getElementById('NowMonth');
            var oNowYear = document.getElementById('NowYear');
            if (!oNowMonth || !oNowYear) { return; }
            oNowYear.value = NowDate.getFullYear();
            oNowMonth.value = NowDate.getMonth() + 1;
            ChangeDate();
        }

        function ChangeDate() {
            //console.log("function ChangeDate()");
            var oNowMonth = document.getElementById('NowMonth');
            var oNowYear = document.getElementById('NowYear');
            if (!oNowMonth || !oNowYear) { return; }

            if (isUnsignedInt(oNowYear.value)) {
                //console.log("oNowYear.value :" + oNowYear.value);
                //console.log("oNowMonth.value:" + oNowMonth.value);
                document.getElementById('Years').innerHTML = oNowYear.value;
                document.getElementById('Months').innerHTML = oNowMonth.value;
                var FirstDate = new Date(oNowYear.value + '/' + oNowMonth.value + '/' + '1')
                var LastDate = new Date(addDateByDay(addDateByMonth(FirstDate, 1), -1))
                var TempDate = new Date(FirstDate);
                //console.log("FirstDate :" + FirstDate);
                //console.log("LastDate :" + LastDate);
                //console.log("TempDate :" + TempDate);
                var MyTable = document.getElementById('DayTable');
                var RowCount = MyTable.rows.length
                /*for (var i = 0; i < RowCount - 1; i++) {  MyTable.deleteRow(); }*/
                for (var i = 0; i < RowCount; i++) {
                    MyTable.deleteRow(0);
                }
                //console.log("for (var i = 0; i < RowCount; i++):");
                var MyRow;
                var MyCell;
                var theDay = FirstDate.getDay();
                if (theDay != 0) {
                    MyRow = MyTable.insertRow();
                    for (var i = 0; i < theDay; i++) {
                        MyCell = MyRow.insertCell();
                        MyCell.innerHTML = '&nbsp;';
                    }
                }
                var i_max_N = 0;
                while (compareDate(TempDate, LastDate) != 1) {
                    i_max_N += 1;
                    //console.log("i_max_N:" + i_max_N);
                    //console.log("LastDate :" + LastDate);
                    //console.log("TempDate :" + TempDate);
                    //console.log("compareDate(TempDate, LastDate):" + compareDate(TempDate, LastDate));
                    if (TempDate.getDay() == 0) { MyRow = MyTable.insertRow(); }
                    MyCell = MyRow.insertCell();
                    MyCell.innerHTML = TempDate.getDate();
                    MyCell.style.textAlign = 'center';
                    MyCell.onmouseover = function () { this.style.backgroundColor = '#F0F0F0' };
                    if (compareDate(TempDate, NowDate) == 0) {
                        MyCell.style.backgroundColor = '#FFFFCC';
                        MyCell.onmouseout = function () { this.style.backgroundColor = '#FFFFCC' };
                    }
                    else {
                        switch (TempDate.getDay()) {
                            case 0:
                                MyCell.style.backgroundColor = '#FFD2D2';
                                MyCell.onmouseout = function () { this.style.backgroundColor = '#FFD2D2' };
                                break;
                            case 6:
                                MyCell.style.backgroundColor = '#CCFFCC';
                                MyCell.onmouseout = function () { this.style.backgroundColor = '#CCFFCC' };
                                break;
                            default:
                                MyCell.onmouseout = function () { this.style.backgroundColor = '' };
                        }
                    }
                    if (compareDate(TempDate, STDate) != -1 && compareDate(TempDate, FTDate) != 1) {
                        MyCell.MyDate = TempDate;
                        MyCell.onclick = returnDate;
                        MyCell.style.cursor = 'hand';
                    }
                    else {
                        MyCell.style.color = '#CCCCCC';
                    }
                    TempDate = new Date(addDateByDay(TempDate, 1));
                    if (i_max_N >= 35) { break; }
                    //console.log("TempDate :" + TempDate);
                    //console.log("compareDate(TempDate, LastDate):" + compareDate(TempDate, LastDate));
                }
            }
        }

        function returnDate(num) {
            var target = undefined;
            if (num) {
                target = num.target;
            } else {
                //firfox不支援
                target = event.srcElement;
            }
            // alert(ValueField);
            if (ValueField == '') {
                //returnValue = ExchangeDate(event.srcElement.MyDate);
                returnValue = ExchangeDate(target.MyDate);
            }
            else {
                //opener.document.getElementById(ValueField).value = ExchangeDate(event.srcElement.MyDate);
                opener.document.getElementById(ValueField).value = ExchangeDate(target.MyDate);
            }
            if (btn != '' && opener.document.getElementById(btn))
                opener.document.getElementById(btn).click();
            self.close();
        }

        //轉出正確的日期格式
        function ExchangeDate(newdate) {
            //console.log("function ExchangeDate(newdate)");
            var result = "";
            if (!newdate) { return result; }
            //console.log("newdate.getFullYear():" + newdate.getFullYear());
            //console.log("(newdate.getMonth() + 1):" + (newdate.getMonth() + 1));
            //console.log("newdate.getDate():" + newdate.getDate());
            result = newdate.getFullYear() + "/" + (newdate.getMonth() + 1) + "/" + newdate.getDate();
            //console.log("result:" + result);
            return result;
        }

        function AddYear(num) {
            //console.log("function AddYear(num)");
            var oNowYear = document.getElementById('NowYear');
            if (!oNowYear) { return; }
            oNowYear.value = parseInt(oNowYear.value, 10) + parseInt(num, 10);
            ChangeDate();
        }
        function AddMonth(num) {
            //console.log("function AddMonth(num)");
            var oNowMonth = document.getElementById('NowMonth');
            var oNowYear = document.getElementById('NowYear');
            if (!oNowMonth || !oNowYear) { return; }
            var NowMonth = parseInt(oNowMonth.value, 10);
            var NewMonth = NowMonth + parseInt(num, 10);
            if (NewMonth > 12) {
                oNowYear.value = parseInt(oNowYear.value, 10) + 1;
                oNowMonth.value = 1;
            }
            else if (NewMonth < 1) {
                oNowYear.value = parseInt(oNowYear.value, 10) - 1;
                oNowMonth.value = 12;
            }
            else {
                oNowMonth.value = NewMonth;
            }
            ChangeDate();
        }
        function ChangeText(obj, underline) {
            //console.log("function ChangeText(obj, underline)");
            obj.style.textDecoration = underline;
        }
    </script>
    <%--<LINK href="../style.css" type="text/css" rel="stylesheet">--%>


    <link href="../css/css.css" rel="stylesheet" type="text/css" />
</head>
<body topmargin="15">
    <form id="form1" method="post" runat="server">
        <font face="新細明體">
            <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="320" border="0">
                <tr>
                    <td>西元<asp:Label ID="Years" runat="server"></asp:Label>年
					    <asp:Label ID="Months" runat="server"></asp:Label>月
                    </td>
                </tr>
                <tr bgcolor="#cbd3e7">
                    <td>
                        <table class="font" id="Table2" style="border-collapse: collapse" cellspacing="0" cellpadding="1" width="100%" border="1">
                            <tr>
                                <td align="center" colspan="4">西元<asp:TextBox ID="NowYear" runat="server" Columns="5"></asp:TextBox>年
                                    <asp:DropDownList ID="NowMonth" runat="server">
                                        <asp:ListItem Value="1">1月</asp:ListItem>
                                        <asp:ListItem Value="2">2月</asp:ListItem>
                                        <asp:ListItem Value="3">3月</asp:ListItem>
                                        <asp:ListItem Value="4">4月</asp:ListItem>
                                        <asp:ListItem Value="5">5月</asp:ListItem>
                                        <asp:ListItem Value="6">6月</asp:ListItem>
                                        <asp:ListItem Value="7">7月</asp:ListItem>
                                        <asp:ListItem Value="8">8月</asp:ListItem>
                                        <asp:ListItem Value="9">9月</asp:ListItem>
                                        <asp:ListItem Value="10">10月</asp:ListItem>
                                        <asp:ListItem Value="11">11月</asp:ListItem>
                                        <asp:ListItem Value="12">12月</asp:ListItem>
                                    </asp:DropDownList>月
                                </td>
                            </tr>
                            <tr>
                                <td onclick="AddYear(-1);" style="cursor: pointer;" onmouseover="ChangeText(this,'underline');" onmouseout="ChangeText(this,'');" align="center">[去年] </td>
                                <td onclick="AddMonth(-1);" style="cursor: pointer" onmouseover="ChangeText(this,'underline');" onmouseout="ChangeText(this,'');" align="center">[上月] </td>
                                <td onclick="AddMonth(1);" style="cursor: pointer" onmouseover="ChangeText(this,'underline');" onmouseout="ChangeText(this,'');" align="center">[下月] </td>
                                <td onclick="AddYear(1);" style="cursor: pointer" onmouseover="ChangeText(this,'underline');" onmouseout="ChangeText(this,'');" align="center">[明年] </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td>
                        <table class="font" id="DayTable" style="border-collapse: collapse" cellspacing="0" cellpadding="1" width="100%" border="1">
                            <tr>
                                <td align="center"><strong>星期日</strong> </td>
                                <td align="center"><strong>星期一</strong> </td>
                                <td align="center"><strong>星期二</strong> </td>
                                <td align="center"><strong>星期三</strong> </td>
                                <td align="center"><strong>星期四</strong> </td>
                                <td align="center"><strong>星期五</strong> </td>
                                <td align="center"><strong>星期六</strong> </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </font>
    </form>
    <script type="text/javascript"> Show_Calendar();</script>
</body>
</html>
