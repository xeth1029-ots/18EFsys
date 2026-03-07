<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="TC_03_ICAP.aspx.vb" Inherits="WDAIIP.TC_03_ICAP" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<%--<!DOCTYPE html>--%>

<%--<html xmlns="http://www.w3.org/1999/xhtml">--%>
<html>
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>iCAP標章證號 </title>
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript"></script>
    <style type="text/css">
        /* 在開啟的新視窗中的 CSS */
        body { overflow-y: auto; /* 啟用垂直滾動 */ overflow-x: hidden; /* 隱藏水平滾動 */ }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div align="center" class="whitecol">
            <table id="Table1" class="table_nw" width="100%" cellpadding="1" cellspacing="1">
                <tr style="display: none">
                    <td class="font" align="center" colspan="4"></td>
                </tr>
                <tr>
                    <td class="table_title" colspan="4">iCAP課程資訊</td>
                </tr>
                <tr>
                    <td class="bluecol_need" width="16%">iCAP標章編號</td>
                    <td class="whitecol" width="34%">
                        <asp:Label ID="labCLASS_ID" runat="server" Text=""></asp:Label></td>
                    <td class="bluecol_need" width="16%">案號</td>
                    <td class="whitecol" width="34%">
                        <asp:Label ID="labCASE_ID" runat="server" Text=""></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">單位名稱</td>
                    <td class="whitecol">
                        <asp:Label ID="labCOMPANY" runat="server" Text=""></asp:Label></td>
                    <td class="bluecol">單位統編</td>
                    <td class="whitecol">
                        <asp:Label ID="labC_ID" runat="server" Text=""></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">課程名稱</td>
                    <td colspan="3" class="whitecol">
                        <asp:Label ID="labCCNAME" runat="server" Text=""></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">訓練課程時數</td>
                    <td class="whitecol">
                        <asp:Label ID="labTRAIN_COURSE_HOURS" runat="server" Text=""></asp:Label></td>
                    <td class="bluecol">職能級別</td>
                    <td class="whitecol">
                        <asp:Label ID="labCLASS_LEVEL" runat="server" Text=""></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">主要對象</td>
                    <td colspan="3" class="whitecol">
                        <asp:Label ID="labTARGET" runat="server" Text=""></asp:Label></td>
                </tr>
                <tr>
                    <td class="bluecol">先備條件</td>
                    <td colspan="3" class="whitecol">
                        <asp:Label ID="labESSENTIAL" runat="server" Text=""></asp:Label></td>
                </tr>
            </table>
        </div>
        <div align="center" class="whitecol">
            <asp:ListView ID="ListView1" runat="server">
                <LayoutTemplate>
                    <%--表頭 Table Header--%>
                    <table id="tbCategory" width="100%" cellpadding="1" cellspacing="1" border="0" runat="server">
                        <%--<tr runat="server"><th runat="server" class="table_title">課程單元</th></tr>--%>
                        <%--決定資料長出來後要放在哪裡，沒加會報的錯誤訊息
                            System.InvalidOperationException: '必須在 ListView 'ListView1' 上指定項目預留位置。
                            請將控制項的 ID 屬性設定為 "itemPlaceholder" 來指定項目預留位置。
                            項目預留位置控制項也必須指定 runat="server"。'--%>
                        <tr runat="server" id="itemPlaceholder" />
                    </table>
                </LayoutTemplate>
                <ItemTemplate>
                    <tr runat="server">
                        <td>
                            <%--<asp:Label ID="lblCustomerID" runat="server" Text='<%#Eval("CustomerID")%>'></asp:Label>Text='<%#Eval("UNIT_SEQ")%>'--%>
                            <table class="table_nw" width="100%" cellpadding="1" cellspacing="1" border="0" runat="server">
                                <tbody>
                                    <tr>
                                        <td class="table_title" colspan="4">
                                            <asp:Label ID="lvlabUNITIT_SEQ" runat="server" Text='<%#Eval("UNIT_SEQ")%>'></asp:Label>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol" width="16%">單元名稱</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Label ID="lvlabUNIT_NAME" runat="server" Text='<%#Eval("UNIT_NAME")%>'></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">應具備之資源與專業學經歷(授課教師)</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Label ID="lvlabTEA_EXP" runat="server" Text='<%#Eval("TEA_EXP")%>'></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">應具備之資源與專業學經歷(協助人員)</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Label ID="lvlabSUP_EXP" runat="server" Text='<%#Eval("SUP_EXP")%>'></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">教材</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Label ID="lvlabTEACH_MATERIALS" runat="server" Text='<%#Eval("TEACH_MATERIALS")%>'></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">教具/設備</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Label ID="lvlabTEACH_EQUIPMENT" runat="server" Text='<%#Eval("TEACH_EQUIPMENT")%>'></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">老師名稱</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Label ID="lvlabACTUAL_TEACHER" runat="server" Text='<%#Eval("ACTUAL_TEACHER")%>'></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">教學方法</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Label ID="lvlabTEACH_METHOD" runat="server" Text='<%#Eval("TEACH_METHOD")%>'></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">評量方式</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Label ID="lvlabEVALUATION_METHOD" runat="server" Text='<%#Eval("EVALUATION_METHOD")%>'></asp:Label></td>
                                    </tr>
                                    <tr>
                                        <td class="bluecol">課程大綱</td>
                                        <td colspan="3" class="whitecol">
                                            <asp:Label ID="lvlabOUTLINE" runat="server" Text='<%#Eval("OUTLINE")%>'></asp:Label></td>
                                    </tr>
                                </tbody>
                            </table>
                        </td>
                    </tr>
                </ItemTemplate>
            </asp:ListView>
        </div>
        <div align="center" class="whitecol">
            <br />
        </div>
        <div align="center" class="whitecol">
            <input type="button" value="關閉視窗" onclick="javascript: window.close();" class="asp_button_M" />
            <br />
            <br />
            <asp:Label ID="lblMsg" runat="server" Text="" ForeColor="Red"></asp:Label>
        </div>
        <div align="center" class="whitecol">
            <br />
        </div>
    </form>
</body>
</html>
