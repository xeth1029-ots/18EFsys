<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_023.aspx.vb" Inherits="WDAIIP.SD_14_023" %>


<!DOCTYPE HTML PUBLIC "-//W3C//Dtd html 4.0 transitional//EN">
<html>
<head>
    <title>結訓證書</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link rel="stylesheet" type="text/css" href="../../css/style.css">
    <script type="text/javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        //Button3					
        function CheckPrint() {
            var OCIDValue1 = document.getElementById('OCIDValue1');
            if (OCIDValue1.value == "") {
                alert('請選擇班級');
                return false;
            }
        }

        function SelectAll(Flag) {
            var MyTable = document.getElementById('DataGrid1');
            for (i = 1; i < MyTable.rows.length; i++) {
                MyTable.rows[i].cells[0].children[0].checked = Flag;
            }
        }

        function ClearData() {
            document.getElementById('TMID1').value = '';
            document.getElementById('OCID1').value = '';
            document.getElementById('TMIDValue1').value = '';
            document.getElementById('OCIDValue1').value = '';
        }

        function choose_class() {
            var RIDValue = document.getElementById('RIDValue');
            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';
            openClass('../02/SD_02_ch.aspx?&RID=' + RIDValue.value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" border="0" cellspacing="1" cellpadding="1" width="100%">
            <tr>
                <td>
                    <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                    <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;表單列印&gt;&gt;結訓證書</asp:Label>
                </td>
            </tr>
        </table>
        <table id="table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
            <tr>
                <td>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" width="20%">訓練機構</td>
                            <td class="whitecol" colspan="3" width="80%">
                                <asp:TextBox ID="center" runat="server" Width="60%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <span id="HistoryList2" style="position: absolute; display: none">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr id="ClassTR" runat="server">
                            <td class="bluecol" width="20%">職類/班別</td>
                            <td class="whitecol" colspan="3" width="80%">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="Button4" type="button" value="清除" name="Button4" runat="server" class="asp_button_M">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 210px">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <%--
                        <tr id="TRPlanPoint28" runat="server">
						    <td class="bluecol" width="100">計畫</td>
						    <td class="whitecol" colspan="3">
							    <asp:RadioButtonList ID="PlanPoint" runat="server" RepeatLayout="Flow" RepeatDirection="Horizontal" CssClass="font">
								    <asp:ListItem Value="1" Selected="True">產業人才投資計畫</asp:ListItem>
								    <asp:ListItem Value="2">提升勞工自主學習計畫</asp:ListItem>
							    </asp:RadioButtonList>
						    </td>
					    </tr>
                        --%>
                        <tr>
                            <td class="bluecol" width="20%">證書編碼</td>
                            <td class="whitecol" colspan="3" width="80%">
                                <%--<asp:TextBox ID="txtCert" runat="server" Width="30%" MaxLength="10" onkeyup="this.value=this.value.replace(/[^\d]/,'')"></asp:TextBox>--%>
                                <asp:TextBox ID="txtCert" runat="server" Width="30%" MaxLength="10"></asp:TextBox>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M" />
                    <asp:Button ID="btnPrint2" runat="server" Text="技檢訓練時數清單" CssClass="asp_Export_M" />
                </td>
            </tr>
        </table>
        <asp:HiddenField ID="hidYears" runat="server" />
        <asp:HiddenField ID="hidOCIDValue" runat="server" />
        <%--<asp:HiddenField ID="hidPCSValue" runat="server" />--%>
        <asp:HiddenField ID="hidorgid" runat="server" />
        <%--<asp:HiddenField ID="hidisPointYN" runat="server" />--%>
    </form>
</body>
</html>
