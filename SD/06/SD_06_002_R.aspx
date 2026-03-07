<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_06_002_R.aspx.vb" Inherits="WDAIIP.SD_06_002_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>SD_06_002_R</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function GETvalue() {
            document.getElementById('Button5').click();
        }

        function choose_class() {
            var RID = document.form1.RIDValue.value;
            openClass('../02/SD_02_ch.aspx?RID=' + RID, 'Class');
        }

        function print() {
            if (document.form1.TPlan.selectedIndex == 0) {
                alert('請選擇訓練計畫');
                return false;
            }
        }

        function showFrame(arg) {
            document.getElementById('FrameObj').style.display = arg;
        }

        function searchcheck() {
            var check1 = '';
            if (document.form1.OCIDValue1.value != '') {
                check1 += '1';
            }
            if (document.form1.cjobValue.value != '') {
                check1 += '1';
            }
            if (document.form1.FTDate1.value != '' || document.form1.FTDate2.value != '') {
                check1 += '1';
            }
            if (check1 == '') {
                alert('請先選擇職類班級、通俗職類或結訓日期區間!');
                return false;
            }
        }
    </script>
    <style type="text/css">
        .rptstl {
            line-height: 22px;
            color: #000000;
            font-size: 13px;
        }
    </style>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;加退保管理&gt;&gt;匯出學員加退保名冊</asp:Label>
                </td>
            </tr>
        </table>
        <table id="myTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tbody>
                <tr>
                    <td>
                        <table id="Table3" cellspacing="1" cellpadding="1" width="100%" class="table_nw">
                            <tbody>
                                <tr>
                                    <td class="bluecol_need" width="20%">訓練機構</td>
                                    <td class="whitecol" width="80%">
                                        <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                        <input id="Button2" type="button" value="..." name="Button2" runat="server" class="asp_button_Mini" />
                                        <input id="RIDValue" type="hidden" name="RIDValue" runat="server" /><br />
                                        <asp:Button ID="Button5" Style="display: none" runat="server" Text="Button5"></asp:Button>
                                        <span id="HistoryList2" style="z-index: 1; position: absolute; display: none" onclick="GETvalue()">
                                            <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol_need" width="20%">訓練計畫</td>
                                    <td class="whitecol" width="80%">
                                        <asp:DropDownList ID="TPlan" runat="server"></asp:DropDownList>
                                        <iframe id="FrameObj" style="position: absolute; background-color: white; width: 310px; display: none; height: 23px; left: 104px" scrolling="no" frameborder="0"></iframe>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol" width="20%">職類/班別</td>
                                    <td class="whitecol" width="80%">
                                        <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                        <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                        <input onclick="choose_class()" type="button" value="..." class="asp_button_Mini" />
                                        <input id="OCIDValue1" style="width: 48px; height: 22px" type="hidden" size="2" name="Hidden2" runat="server" />
                                        <input id="TMIDValue1" type="hidden" name="Hidden1" runat="server" /><br />
                                        <span id="HistoryList" style="position: absolute; display: none; left: 30%">
                                            <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                        </span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol" width="20%"><asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label></td>
                                    <td class="whitecol" width="80%">
                                        <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30" Width="30%"></asp:TextBox>
                                        <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="asp_button_Mini" />
                                        <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol" width="20%">結訓日期</td>
                                    <td class="whitecol" width="80%">
                                        <asp:TextBox ID="FTDate1" runat="server" Width="16%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= FTDate1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                        ～<asp:TextBox ID="FTDate2" runat="server" Width="16%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= FTDate2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol" width="20%">加保日期</td>
                                    <td class="whitecol" width="80%">
                                        <asp:TextBox ID="ApplyD1" runat="server" Width="16%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= ApplyD1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                        ～<asp:TextBox ID="ApplyD2" runat="server" Width="16%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= ApplyD2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                    </td>
                                </tr>
                                <tr>
                                    <td class="bluecol" width="20%">退保日期</td>
                                    <td class="whitecol" width="80%">
                                        <asp:TextBox ID="DropoutD1" runat="server" Width="16%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= DropoutD1.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                        ～<asp:TextBox ID="DropoutD2" runat="server" Width="16%"></asp:TextBox>
                                        <span runat="server"><img style="cursor: pointer" onclick="javascript:show_calendar('<%= DropoutD2.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" /></span>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                        <table width="100%">
                            <tr>
                                <td class="whitecol" align="center">
                                    <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
                                    <asp:Button ID="btnExport1" runat="server" Text="匯出加保Excel" CssClass="asp_Export_M"></asp:Button>
                                    <asp:Button ID="btnExport2" runat="server" Text="匯出退保Excel" CssClass="asp_Export_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                        <div>
                            <center><asp:Label ID="lblMsg" runat="server" ForeColor="Red"></asp:Label></center>
                        </div>
                        <div id="Div1" align="center" runat="server">
                            <asp:Table ID="tbRpt" CssClass="rptstl" CellPadding="0" CellSpacing="0" runat="server"></asp:Table>
                        </div>
                    </td>
                </tr>
            </tbody>
        </table>
    </form>
</body>
</html>