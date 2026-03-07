<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_14_008.aspx.vb" Inherits="WDAIIP.SD_14_008" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>訓練課程開班學員名冊</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function printkind() {
            document.getElementById('print_orderyby').value = 'c.IDNO'
            if (getValue("print_type") == '2') {
                document.getElementById('print_orderyby').value = 'a.StudentID'
            }
        }

        function GETvalue() {
            document.getElementById('Button3').click();
        }

        function CheckPrint(filename) {
            return false;
            /*  56509922CDD9A1C7C85F894E39C6ECD741219D3FDF1038D4ED899BA5BF1357146D9484236F70A069E265BAA809E1CF3C37784AA4CA319AC1311B03FB
847A1B635F96168348FD66AADE835754AD606D509B3CD1A9033D99AC87262DD1857984057AB2A0649C187B684435DE9ACEB89ACCD41201B5F4A06126
D57A59F9BE7803A8134613674627ADFE86372E80070EDD52411BCE053D53594741A3D4DE0EEE50687E39D194F2FEF128364A114B1D6518DBE7E4D4DD
9E281920DCBAFF5334557C4432819A18ECDE8FC7A0F31B2446D0B5E35578EFAA838F644C4602469E7F2DB9AEE12A5713246C774B3B0BB4B39C2FB651
53DA758B05ADE4D9A326AC5EA34E0317287DF174C7DE09E99756F816B82C31C88EA162D5A282229EF125511C0743C39E12419ECC21790247689026E4
7B60F858672E3394EDA891A59A47B9A9883615875701410D5765B89673FE6889E3D62D51A9AD3D9E397D40E97A3F26E1343B9E2F59580B0E6B081D98
85F6FB999565A67C46C8427FC4803678ED6006D31355CDF15D4D79280EFDC93F0A9594B04747964979558343FCB508242929AC468BE7CAF6EA08E15C
FEFFE2D7BF96A399651499221F3AADE81F04527A83C3BE066A9A1809DE715759A2D01B415D0D174D83C3C8238D5C5AB484931E499CBAE8C62E372381
3856565A628862480E4679D8DD3EDBE4605BAE43B9D4D09AF621C72C009F32E2E20B3CED2AED2C2F0542D2D9E5B6CBEE538BAA0B4A3AF5009DF0CD01
422FF52519CF2F622170A8033228DB5603B6457C05B8133735BDE7952AAC9FFEF343CA563AFE14EDA090D10978B7954A2958A3C862EFDCF720B389AA
DFEC9110B5AF39FEA4E7DAEF720B6D52557D710713DE570C32DBB520713E84E256745D13A9F6BE71 */
        }

        function choose_class() {
            document.getElementById('OCID1').value = '';
            document.getElementById('TMID1').value = '';
            document.getElementById('OCIDValue1').value = '';
            document.getElementById('TMIDValue1').value = '';
            openClass('../02/SD_02_ch.aspx?&RID=' + document.getElementById('RIDValue').value);
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label>
                                <asp:Label ID="TitleLab2" runat="server">首頁&gt;&gt;學員動態管理&gt;&gt;產學訓表單列印&gt;&gt;訓練課程開班學員名冊</asp:Label>
                            </td>
                        </tr>
                    </table>
                    <table class="table_sch">
                        <tr>
                            <td class="bluecol" style="width: 20%">訓練機構</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="center" runat="server" Width="55%"></asp:TextBox>
                                <input id="RIDValue" type="hidden" name="Hidden2" runat="server">
                                <input id="Button2" type="button" value="..." name="Button2" runat="server" class="button_b_Mini">
                                <asp:Button ID="Button3" Style="display: none" runat="server" Text="Button3"></asp:Button>
                                <span id="HistoryList2" style="position: absolute; display: none" onclick="GETvalue()">
                                    <asp:Table ID="HistoryRID" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol_need">職類/班別</td>
                            <td class="whitecol" colspan="3">
                                <asp:TextBox ID="TMID1" runat="server" onfocus="this.blur()" Width="25%"></asp:TextBox>
                                <asp:TextBox ID="OCID1" runat="server" onfocus="this.blur()" Width="30%"></asp:TextBox>
                                <input onclick="choose_class()" type="button" value="..." class="button_b_Mini">
                                <input id="TMIDValue1" type="hidden" name="TMIDValue1" runat="server">
                                <input id="OCIDValue1" type="hidden" name="OCIDValue1" runat="server">
                                <span id="HistoryList" style="position: absolute; display: none; left: 28%">
                                    <asp:Table ID="HistoryTable" runat="server" Width="100%"></asp:Table>
                                </span>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">列印排序方式</td>
                            <td class="whitecol" colspan="3">
                                <asp:RadioButtonList ID="print_type" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                    <asp:ListItem Value="1">依【身分證字號】排序列印</asp:ListItem>
                                    <asp:ListItem Value="2" Selected="True">依【學號】排序列印</asp:ListItem>
                                </asp:RadioButtonList>
                            </td>
                        </tr>
                        <tr id="tr_ddl_INQUIRY_S" runat="server">
                            <td class="bluecol_need">查詢原因</td>
                            <td class="whitecol" colspan="3">
                                <asp:DropDownList ID="ddl_INQUIRY_Sch" runat="server"></asp:DropDownList>
                            </td>
                        </tr>
                    </table>
                    <div align="center" class="whitecol">
                        <%--<input id="button5" type="button" value="列印" runat="server" class="asp_Export_M">--%>
                        <asp:Button ID="btnPrint1" runat="server" Text="列印" CssClass="asp_Export_M" />
                    </div>
                </td>
            </tr>
        </table>
        <input id="Years" type="hidden" name="Years" runat="server" />
        <input id="print_orderyby" type="hidden" name="print_orderyby" runat="server" />
    </form>
</body>
</html>
