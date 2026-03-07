<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_05_013_R.aspx.vb" Inherits="WDAIIP.SD_05_013_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>列印學員郵遞標籤</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        function choose_class() {
            //var RID = document.form1.RIDValue.value;
            var RIDValue = document.getElementById('RIDValue');
            openClass('../02/SD_02_ch.aspx?special=2&RID=' + RIDValue);
        }

        function search() {
            var msg = '';
            if (isEmpty(document.form1.STDate1) && isEmpty(document.form1.STDate2) && isEmpty(document.form1.FTDate1) &&
				isEmpty(document.form1.FTDate2) && isEmpty(document.form1.OCID1)) {
                msg += '查詢條件請選擇其中之一!\n';
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
    <table id="FrameTable" cellspacing="1" cellpadding="1" width="740" border="0">
        <tr>
            <td align="center">
                <table class="font" id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務管理&gt;&gt;<font color="#990000">列印學員郵遞標籤</font>
                        </td>
                    </tr>
                </table>
                <table class="table_sch" id="Table3" cellpadding="1" cellspacing="1">
                    <tr>
                        <td width="100" class="bluecol">
                            開訓日期區間
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="STDate1" runat="server" Width="100px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~
                            <asp:TextBox ID="STDate2" runat="server" Width="100px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= STDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        </td>
                    </tr>
                    <tr>
                        <td width="100" class="bluecol">
                            結訓日期區間
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="FTDate1" runat="server" Width="100px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= FTDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">~
                            <asp:TextBox ID="FTDate2" runat="server" Width="100px"></asp:TextBox><img style="cursor: pointer" onclick="javascript:show_calendar('<%= FTDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30">
                        </td>
                    </tr>
                    <tr>
                        <td width="100" class="bluecol">
                            職類/班別
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="210px"></asp:TextBox>
                            <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="200px"></asp:TextBox>
                            <input onclick="choose_class();" type="button" value="..." />
                            <input id="RIDValue" type="hidden" name="RIDValue" runat="server" />
                            <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server" />
                            <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server" />
                            <span id="HistoryList" style="display: none; left: 270px; position: absolute">
                                <asp:Table ID="HistoryTable" runat="server" Width="310">
                                </asp:Table>
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            通俗職類
                        </td>
                        <td class="whitecol" colspan="3">
                            <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30"></asp:TextBox>
                            <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server">
                            <input id="cjobValue" type="hidden" name="cjobValue" runat="server">
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            學員狀態
                        </td>
                        <td class="whitecol">
                            <asp:RadioButtonList ID="IsOnJob" runat="server" RepeatDirection="Horizontal" CssClass="font" RepeatLayout="Flow">
                                <asp:ListItem Value="0">未就業</asp:ListItem>
                                <asp:ListItem Value="1">已就業</asp:ListItem>
                            </asp:RadioButtonList>
                        </td>
                    </tr>
                </table>
                <table width="100%">
                    <tr>
                        <td colspan="2" class="whitecol">
                            <p align="center">
                                &nbsp;<asp:Button ID="Button2" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                            </p>
                        </td>
                    </tr>
                </table>
                <asp:Label ID="msg" runat="server" CssClass="font" ForeColor="Red"></asp:Label>
            </td>
        </tr>
    </table>
    </form>
</body>
</html>
