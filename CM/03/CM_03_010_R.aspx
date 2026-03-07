<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="CM_03_010_R.aspx.vb" Inherits="WDAIIP.CM_03_010_R" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>列印報表</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <style type="text/css">
        @page :left { margin-left: 4cm; margin-right: 3cm; }
        @page :right { margin-left: 3cm; margin-right: 4cm; }
    </style>
    <!-- MeadCo ScriptX -->
    <%--<object id="factory" style="display: none" codebase="../../scriptx/smsx.cab#Version=6,6,440,26" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" viewastext></object>--%>
    <script type="text/javascript" language="javascript">
        var isprint = false;

        function Show_Background() {
            document.getElementById('tb_01').style.display = "none";
            document.frames[0].ShowFrame(document.body.innerHTML);
        }

        function Set_FrameHeight(hh) {
            document.body.innerHTML = document.body.innerHTML + "<IMG src='../../images/rptpic/temple/TIMS_2.jpg' style='z-index:-1;position:absolute;top:0px;left:0px;height:1500px;width:1080px;display:inline' />";
            Auto_Print();
            document.getElementById('tb_01').style.display = "inline";
        }

        function Auto_Print() {
            window.print();
            //if (!factory.object) {
            //    return
            //} else {
            //    document.all.factory.printing.header = ""; //頁首，空白為不印頁首，也就不會佔空間
            //    document.all.factory.printing.footer = ""; //註腳，空白為不印註腳，也就不會佔空間
            //    document.all.factory.printing.leftMargin = 6; //左邊界
            //    document.all.factory.printing.topMargin = 0; //上邊界
            //    document.all.factory.printing.rightMargin = 0; //右邊界
            //    document.all.factory.printing.bottomMargin = 0; //下邊界
            //    document.all.factory.printing.portrait = false; //直印，false:橫印 
            //    document.all.factory.printing.Print(true);
            //    isprint = true;
            //}
        }

        function bufferTime() {
            setTimeout("goBack()", 500);
        }

        function goBack() {
            if (isprint = true) {
                isprint = false;
                bufferTime();
            } else {
                window.history.go(0);
            }
        }
    </script>
</head>
<body background="../../images/rptpic/temple/TIMS_1.jpg">
    <form id="form1" method="post" runat="server">
        <table id="tb_01" align="left">
            <tr>
                <td align="left">
                    <img id="printRpt" style="cursor: pointer" onclick="if(!isprint){ Show_Background();}else{Auto_Print();}goBack();" alt="列印報表" src="../../images/rptpic/Print.gif">&nbsp;&nbsp;
				<asp:ImageButton ID="bt_excel" Style="cursor: pointer" runat="server" ImageUrl="../../images/rptpic/Excel.gif" AlternateText="匯出Excel"></asp:ImageButton>&nbsp; </td>
            </tr>
        </table>
        <br style="line-height: 50px">
        <div id="div1" style="width: 1070px" align="center" runat="server">
            <span style="font-weight: bold; font-size: 18px">職業訓練各類身分別開訓人數統計</span>
            <table style="font-size: 14px" width="100%">
                <tr id="trYear_1" runat="server">
                    <td>計畫年度：<asp:Label ID="labYear_1" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr id="trSDate_1" runat="server">
                    <td>開訓期間：<asp:Label ID="labSDate_1" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr id="trEDate_1" runat="server">
                    <td>結訓期間：<asp:Label ID="labEDate_1" runat="server"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td>統計縣市：<asp:Label ID="labCityName_1" runat="server"></asp:Label>
                    </td>
                </tr>
            </table>
            <table style="font-size: 14px; border-collapse: collapse" bordercolor="black" cellspacing="0" cellpadding="0" width="100%" border="1">
                <tr>
                    <td style="font-weight: bold" align="center" width="100">開訓人數 </td>
                    <td>男性：<asp:Label ID="labM_1" runat="server" Width="50px"></asp:Label>人,&nbsp; 女性：<asp:Label ID="labF_1" runat="server" Width="50px"></asp:Label>人,&nbsp; 合計：<asp:Label ID="labTotal_1" runat="server" Width="50px"></asp:Label>人 </td>
                </tr>
                <tr>
                    <td style="font-weight: bold" align="center">開訓預算別 </td>
                    <td>公務：<asp:Label ID="labBudget01_1" runat="server" Width="50px"></asp:Label>人,&nbsp; 就安：<asp:Label ID="labBudget02_1" runat="server" Width="50px"></asp:Label>人,&nbsp; 就保：<asp:Label ID="labBudget03_1" runat="server" Width="50px"></asp:Label>人,&nbsp; 特別預算：<asp:Label ID="labBudget98_1" runat="server" Width="50px"></asp:Label>人,&nbsp; 不補助：<asp:Label ID="labBudget99_1" runat="server" Width="50px"></asp:Label>人,&nbsp; 其他：<asp:Label ID="labBudgetOdr_1" runat="server" Width="50px"></asp:Label>人,&nbsp; 合計：<asp:Label ID="labBudgetTotal_1" runat="server" Width="50px"></asp:Label>人 </td>
                </tr>
                <tr>
                    <td style="font-weight: bold" valign="middle" align="center">開訓特定對象<br>
                        辦理人數統計 </td>
                    <td>
                        <table style="font-size: 12px" cellspacing="5" cellpadding="5">
                            <tr>
                                <td width="18%">01.一般身分者 </td>
                                <td width="2%">
                                    <asp:Label ID="labId01_1" runat="server"></asp:Label>
                                </td>
                                <td width="18%">02.就業保險被保險人非自願失業者 </td>
                                <td width="2%">
                                    <asp:Label ID="labId02_1" runat="server"></asp:Label>
                                </td>
                                <td width="18%">03.負擔家計婦女 </td>
                                <td width="2%">
                                    <asp:Label ID="labId03_1" runat="server"></asp:Label>
                                </td>
                                <td width="18%">04.中高齡者 </td>
                                <td width="2%">
                                    <asp:Label ID="labId04_1" runat="server"></asp:Label>
                                </td>
                                <td width="18%">05.原住民 </td>
                                <td width="2%">
                                    <asp:Label ID="labId05_1" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>06.身心障礙者 </td>
                                <td>
                                    <asp:Label ID="labId06_1" runat="server"></asp:Label>
                                </td>
                                <td>07.生活扶助戶 </td>
                                <td>
                                    <asp:Label ID="labId07_1" runat="server"></asp:Label>
                                </td>
                                <td>08.急難救助戶 </td>
                                <td>
                                    <asp:Label ID="labId08_1" runat="server"></asp:Label>
                                </td>
                                <td>09.家庭暴力受害人 </td>
                                <td>
                                    <asp:Label ID="labId09_1" runat="server"></asp:Label>
                                </td>
                                <td>10.更生受保護人 </td>
                                <td>
                                    <asp:Label ID="labId10_1" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>11.農漁民 </td>
                                <td>
                                    <asp:Label ID="labId11_1" runat="server"></asp:Label>
                                </td>
                                <td>12.屆退官兵(須單位將級以上長官薦送函) </td>
                                <td>
                                    <asp:Label ID="labId12_1" runat="server"></asp:Label>
                                </td>
                                <td>13.外籍配偶 </td>
                                <td>
                                    <asp:Label ID="labId13_1" runat="server"></asp:Label>
                                </td>
                                <td>14.大陸配偶 </td>
                                <td>
                                    <asp:Label ID="labId14_1" runat="server"></asp:Label>
                                </td>
                                <td>15.遊民 </td>
                                <td>
                                    <asp:Label ID="labId15_1" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>16.公營事業民營化員工 </td>
                                <td>
                                    <asp:Label ID="labId16_1" runat="server"></asp:Label>
                                </td>
                                <td>17.參加職業工會失業者 </td>
                                <td>
                                    <asp:Label ID="labId17_1" runat="server"></asp:Label>
                                </td>
                                <td>18.921受災戶 </td>
                                <td>
                                    <asp:Label ID="labId18_1" runat="server"></asp:Label>
                                </td>
                                <td>19.性侵害被害人 </td>
                                <td>
                                    <asp:Label ID="labId19_1" runat="server"></asp:Label>
                                </td>
                                <td>20.就業保險被保險人自願失業者 </td>
                                <td>
                                    <asp:Label ID="labId20_1" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>21.臨時工作津貼人員 </td>
                                <td>
                                    <asp:Label ID="labId21_1" runat="server"></asp:Label>
                                </td>
                                <td>22.多元就業開發方案人員 </td>
                                <td>
                                    <asp:Label ID="labId22_1" runat="server"></asp:Label>
                                </td>
                                <td>23.申請失業給付經失業認定者(學習卷專用) </td>
                                <td>
                                    <asp:Label ID="labId23_1" runat="server"></asp:Label>
                                </td>
                                <td>24.非失業認定之就業保險失業者(學習卷專用) </td>
                                <td>
                                    <asp:Label ID="labId24_1" runat="server"></asp:Label>
                                </td>
                                <td>25.非就業保險失業者(學習卷專用) </td>
                                <td>
                                    <asp:Label ID="labId25_1" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>26.犯罪被害人及其親屬 </td>
                                <td>
                                    <asp:Label ID="labId26_1" runat="server"></asp:Label>
                                </td>
                                <td>27.長期失業者 </td>
                                <td>
                                    <asp:Label ID="labId27_1" runat="server"></asp:Label>
                                </td>
                                <td>28.獨力負擔家計者 </td>
                                <td>
                                    <asp:Label ID="labId28_1" runat="server"></asp:Label>
                                </td>
                                <td>29.天然災害受災民眾 </td>
                                <td>
                                    <asp:Label ID="labId29_1" runat="server"></asp:Label>
                                </td>
                                <td>30.因應貿易自由化協助勞工 </td>
                                <td>
                                    <asp:Label ID="labId30_1" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>31.單一中華民國國籍之無戶籍國民 </td>
                                <td>
                                    <asp:Label ID="labId31_1" runat="server"></asp:Label>
                                </td>
                                <td>32.取得居留身分之泰國、緬甸、印度或尼泊爾地區無國籍人民 </td>
                                <td>
                                    <asp:Label ID="labId32_1" runat="server"></asp:Label>
                                </td>
                                <td>99.其他 </td>
                                <td>
                                    <asp:Label ID="labIdOdr_1" runat="server"></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
            <br style='margin: 0cm 0cm 0pt; line-height: 80px; mso-pagination: widow-orphan'>
            <br clear="all" style='page-break-before: always; mso-special-character: line-break'>
            <br style="line-height: 50px">
            <span style="font-weight: bold; font-size: 18px">職業訓練各類身分別結訓人數統計</span>
            <table style="font-size: 14px" width="100%">
                <tr id="trYear_2" runat="server">
                    <td>計畫年度：<asp:Label ID="labYear_2" runat="server"></asp:Label></td>
                </tr>
                <tr id="trSDate_2" runat="server">
                    <td>開訓期間：<asp:Label ID="labSDate_2" runat="server"></asp:Label></td>
                </tr>
                <tr id="trEDate_2" runat="server">
                    <td>結訓期間：<asp:Label ID="labEDate_2" runat="server"></asp:Label></td>
                </tr>
                <tr>
                    <td>統計縣市：<asp:Label ID="labCityName_2" runat="server"></asp:Label></td>
                </tr>
            </table>
            <table style="font-size: 14px; border-collapse: collapse" bordercolor="black" cellspacing="0" cellpadding="0" width="100%" border="1">
                <tr>
                    <td style="font-weight: bold" align="center" width="100">結訓人數 </td>
                    <td>男性：<asp:Label ID="labM_2" Width="50px" runat="server"></asp:Label>人,&nbsp; 女性：<asp:Label ID="labF_2" Width="50px" runat="server"></asp:Label>人,&nbsp; 合計：<asp:Label ID="labTotal_2" Width="50px" runat="server"></asp:Label>人 </td>
                </tr>
                <tr>
                    <td style="font-weight: bold" align="center">結訓預算別 </td>
                    <td>公務：<asp:Label ID="labBudget01_2" Width="50px" runat="server"></asp:Label>人,&nbsp; 就安：<asp:Label ID="labBudget02_2" Width="50px" runat="server"></asp:Label>人,&nbsp; 就保：<asp:Label ID="labBudget03_2" Width="50px" runat="server"></asp:Label>人,&nbsp; 特別預算：<asp:Label ID="labBudget98_2" Width="50px" runat="server"></asp:Label>人,&nbsp; 不補助：<asp:Label ID="labBudget99_2" Width="50px" runat="server"></asp:Label>人,&nbsp; 其他：<asp:Label ID="labBudgetOdr_2" runat="server" Width="50px"></asp:Label>人,&nbsp; 合計：<asp:Label ID="labBudgetTotal_2" Width="50px" runat="server"></asp:Label>人 </td>
                </tr>
                <tr>
                    <td style="font-weight: bold" valign="middle" align="center">結訓特定對象<br>
                        辦理人數統計 </td>
                    <td>
                        <table style="font-size: 12px" cellspacing="5" cellpadding="5">
                            <tr>
                                <td width="18%">01.一般身分者 </td>
                                <td width="2%">
                                    <asp:Label ID="labId01_2" runat="server"></asp:Label>
                                </td>
                                <td width="18%">02.就業保險被保險人非自願失業者 </td>
                                <td width="2%">
                                    <asp:Label ID="labId02_2" runat="server"></asp:Label>
                                </td>
                                <td width="18%">03.負擔家計婦女 </td>
                                <td width="2%">
                                    <asp:Label ID="labId03_2" runat="server"></asp:Label>
                                </td>
                                <td width="18%">04.中高齡者 </td>
                                <td width="2%">
                                    <asp:Label ID="labId04_2" runat="server"></asp:Label>
                                </td>
                                <td width="18%">05.原住民 </td>
                                <td width="2%">
                                    <asp:Label ID="labId05_2" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>06.身心障礙者 </td>
                                <td>
                                    <asp:Label ID="labId06_2" runat="server"></asp:Label>
                                </td>
                                <td>07.生活扶助戶 </td>
                                <td>
                                    <asp:Label ID="labId07_2" runat="server"></asp:Label>
                                </td>
                                <td>08.急難救助戶 </td>
                                <td>
                                    <asp:Label ID="labId08_2" runat="server"></asp:Label>
                                </td>
                                <td>09.家庭暴力受害人 </td>
                                <td>
                                    <asp:Label ID="labId09_2" runat="server"></asp:Label>
                                </td>
                                <td>10.更生受保護人 </td>
                                <td>
                                    <asp:Label ID="labId10_2" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>11.農漁民 </td>
                                <td>
                                    <asp:Label ID="labId11_2" runat="server"></asp:Label>
                                </td>
                                <td>12.屆退官兵(須單位將級以上長官薦送函) </td>
                                <td>
                                    <asp:Label ID="labId12_2" runat="server"></asp:Label>
                                </td>
                                <td>13.外籍配偶 </td>
                                <td>
                                    <asp:Label ID="labId13_2" runat="server"></asp:Label>
                                </td>
                                <td>14.大陸配偶 </td>
                                <td>
                                    <asp:Label ID="labId14_2" runat="server"></asp:Label>
                                </td>
                                <td>15.遊民 </td>
                                <td>
                                    <asp:Label ID="labId15_2" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>16.公營事業民營化員工 </td>
                                <td>
                                    <asp:Label ID="labId16_2" runat="server"></asp:Label>
                                </td>
                                <td>17.參加職業工會失業者 </td>
                                <td>
                                    <asp:Label ID="labId17_2" runat="server"></asp:Label>
                                </td>
                                <td>18.921受災戶 </td>
                                <td>
                                    <asp:Label ID="labId18_2" runat="server"></asp:Label>
                                </td>
                                <td>19.性侵害被害人 </td>
                                <td>
                                    <asp:Label ID="labId19_2" runat="server"></asp:Label>
                                </td>
                                <td>20.就業保險被保險人自願失業者 </td>
                                <td>
                                    <asp:Label ID="labId20_2" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>21.臨時工作津貼人員 </td>
                                <td>
                                    <asp:Label ID="labId21_2" runat="server"></asp:Label>
                                </td>
                                <td>22.多元就業開發方案人員 </td>
                                <td>
                                    <asp:Label ID="labId22_2" runat="server"></asp:Label>
                                </td>
                                <td>23.申請失業給付經失業認定者(學習卷專用) </td>
                                <td>
                                    <asp:Label ID="labId23_2" runat="server"></asp:Label>
                                </td>
                                <td>24.非失業認定之就業保險失業者(學習卷專用) </td>
                                <td>
                                    <asp:Label ID="labId24_2" runat="server"></asp:Label>
                                </td>
                                <td>25.非就業保險失業者(學習卷專用) </td>
                                <td>
                                    <asp:Label ID="labId25_2" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>26.犯罪被害人及其親屬 </td>
                                <td>
                                    <asp:Label ID="labId26_2" runat="server"></asp:Label>
                                </td>
                                <td>27.長期失業者 </td>
                                <td>
                                    <asp:Label ID="labId27_2" runat="server"></asp:Label>
                                </td>
                                <td>28.獨力負擔家計者 </td>
                                <td>
                                    <asp:Label ID="labId28_2" runat="server"></asp:Label>
                                </td>
                                <td>29.天然災害受災民眾 </td>
                                <td>
                                    <asp:Label ID="labId29_2" runat="server"></asp:Label>
                                </td>
                                <td>30.因應貿易自由化協助勞工 </td>
                                <td>
                                    <asp:Label ID="labId30_2" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>31.單一中華民國國籍之無戶籍國民 </td>
                                <td>
                                    <asp:Label ID="labId31_2" runat="server"></asp:Label>
                                </td>
                                <td>32.取得居留身分之泰國、緬甸、印度或尼泊爾地區無國籍人民 </td>
                                <td>
                                    <asp:Label ID="labId32_2" runat="server"></asp:Label>
                                </td>
                                <td>99.其他 </td>
                                <td>
                                    <asp:Label ID="labIdOdr_2" runat="server"></asp:Label>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
        <iframe id="ifram1" src="../../RPT.htm" width="100%" height="0"></iframe>
    </form>
</body>
</html>
