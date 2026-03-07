<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_10_MBR.aspx.vb" Inherits="WDAIIP.TC_10_MBR" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>請選擇審查委員名單</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <%--<LINK href="../style.css" type="text/css" rel="stylesheet">--%>
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        function ReturnMBR(vEMSEQ, vMBRNAME) {
            opener.document.getElementById(getParamValue('ValueField')).value = vEMSEQ;
            opener.document.getElementById(getParamValue('TextField')).value = vMBRNAME;
            window.close();
        }

        function ReturnMBR2() {
            opener.document.getElementById(getParamValue('ValueField')).value = document.getElementById('VAL_EMSEQ').value;
            opener.document.getElementById(getParamValue('TextField')).value = document.getElementById('VAL_MBRNAME').value;
            window.close();
        }

        function SelectMBR(Flag, vEMSEQ, vMBRNAME) {
            var VAL_EMSEQ = document.getElementById('VAL_EMSEQ');
            var VAL_MBRNAME = document.getElementById('VAL_MBRNAME');
            if (Flag) {
                //add
                if (VAL_EMSEQ.value != '') { VAL_EMSEQ.value += ','; }
                VAL_EMSEQ.value += vEMSEQ;
                if (VAL_MBRNAME.value != '') { VAL_MBRNAME.value += ','; }
                VAL_MBRNAME.value += vMBRNAME;
            }
            else {
                //del
                if (VAL_EMSEQ.value.indexOf(',' + vEMSEQ + ',') != -1) {
                    VAL_EMSEQ.value = VAL_EMSEQ.value.replace(',' + vEMSEQ, '')
                    VAL_MBRNAME.value = VAL_MBRNAME.value.replace(',' + vMBRNAME, '')
                }
                else if (VAL_EMSEQ.value.indexOf(',' + vEMSEQ) != -1) {
                    VAL_EMSEQ.value = VAL_EMSEQ.value.replace(',' + vEMSEQ, '')
                    VAL_MBRNAME.value = VAL_MBRNAME.value.replace(',' + vMBRNAME, '')
                }
                else if (VAL_EMSEQ.value.indexOf(vEMSEQ + ',') != -1) {
                    VAL_EMSEQ.value = VAL_EMSEQ.value.replace(vEMSEQ + ',', '')
                    VAL_MBRNAME.value = VAL_MBRNAME.value.replace(vMBRNAME + ',', '')
                }
                else if (VAL_EMSEQ.value.indexOf(vEMSEQ) != -1) {
                    VAL_EMSEQ.value = VAL_EMSEQ.value.replace(vEMSEQ, '')
                    VAL_MBRNAME.value = VAL_MBRNAME.value.replace(vMBRNAME, '')
                }
            }
        }
    </script>

</head>
<body>
    <form id="form1" method="post" runat="server">
        <table class="font" id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr id="ProTR1" runat="server">
                <td>
                    <table class="table_nw" id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td class="bluecol" width="20%">審查委員姓名： </td>
                            <td class="whitecol" width="80%">
                                <asp:TextBox ID="MBRNAME" runat="server"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td align="center" colspan="2" class="whitecol">
                                <asp:Button ID="Button1" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                <input onclick="ReturnMBR('', '')" type="button" value="清除" class="asp_button_M" />
                            </td>
                        </tr>
                    </table>

                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <div style="overflow-y: auto; height: 550px;">
                        <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" CssClass="font" AutoGenerateColumns="False" CellPadding="8">
                            <AlternatingItemStyle BackColor="WhiteSmoke" />
                            <HeaderStyle HorizontalAlign="Center" CssClass="head_navy" />
                            <Columns>
                                <asp:TemplateColumn HeaderStyle-Width="10%">
                                    <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    <ItemTemplate><input id="Checkbox1" type="checkbox" runat="server" /></ItemTemplate>
                                </asp:TemplateColumn>
                                <asp:BoundColumn DataField="RECRUIT_N" HeaderText="遴聘類別"></asp:BoundColumn>
                                <asp:BoundColumn DataField="MBRNAME" HeaderText="審查委員姓名"></asp:BoundColumn>
                                <asp:BoundColumn DataField="UNITNAME" HeaderText="現職服務機構"></asp:BoundColumn>
                                <asp:BoundColumn DataField="JOBTITLE" HeaderText="職稱"></asp:BoundColumn>
                                <%--<asp:BoundColumn DataField="PUSHDISTID" HeaderText="推薦分署"></asp:BoundColumn>--%>
                                <asp:TemplateColumn HeaderText="推薦分署">
                                    <ItemTemplate>
                                        <asp:Label ID="labPUSHDISTID_N" runat="server"></asp:Label>
                                    </ItemTemplate>
                                </asp:TemplateColumn>
                            </Columns>
                        </asp:DataGrid>
                    </div>
                    <input id="Button2" type="button" value="送出" name="Button2" runat="server" onclick="ReturnMBR2();" class="asp_button_M" />
                </td>
            </tr>
        </table>

        <input id="VAL_EMSEQ" runat="server" type="hidden" />
        <input id="VAL_MBRNAME" runat="server" type="hidden" />

        <%--        <input id="VAL_EMSEQ" type="hidden" name="TeachID" runat="server" />
        <input id="VAL_MBRNAME" type="hidden" name="TeachName" runat="server" />--%>
    </form>
</body>
</html>
