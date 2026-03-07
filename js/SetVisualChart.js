function formatFloat(num, pos) {
    var size = Math.pow(10, pos);
    return Math.round(num * size) / size;
}

function formatNumber(n) {
    n += "";
    var arr = n.split(".");
    var re = /(\d{1,3})(?=(\d{3})+$)/g;
    return arr[0].replace(re, "$1,") + (arr.length == 2 ? "." + arr[1] : "");
}

//自辦在職：指標1_各分署辦理訓練人次統計
/*
   categories: [
                '北基宜花金馬分署_在職進修訓練',
                '北基宜花金馬分署_接受企業委託訓練',
                '桃竹苗分署__在職進修訓練',
                '桃竹苗分署_接受企業委託訓練',
                '中彰投分署_在職進修訓練',
                '中彰投分署_接受企業委託訓練',
                '雲嘉南分署_在職進修訓練',
                '雲嘉南分署_接受企業委託訓練',
                '高屏澎東分署_在職進修訓練',
                '高屏澎東分署_接受企業委託訓練'
            ]
*/
function creatChartB_1(containerID, data1, data2, data3, data4, data5) {
    $('#' + containerID).highcharts({
        chart: {
            type: 'column'
        },
        title: {
            text: '各分署辦理訓練人次統計',
            fontsize: 12
        },
        xAxis: {
            categories: data5
        },
        yAxis: [{ // Primary yAxis
            labels: {
                formatter: function () {
                    return this.axis.defaultLabelFormatter.call(this);
                }
            },
            title: {
                text: '人數'
            }
        }, { // Secondary yAxis
            title: {
                text: '百分比'
            },
            opposite: true,
            labels: {
                formatter: function () {
                    return formatFloat(this.value * 100, 2) + '%';
                }
            },
            min: 0,
            max: 1
        }],
        legend: {
            shadow: false
        },
        tooltip: {
            formatter: function () {
                var s = '<b>' + this.x + '</b>';
                $.each(this.points, function () {
                    if (this.series.name.indexOf('率') > 0) {
                        s += '<br/>' + this.series.name + ': ' +
                            formatFloat(this.y * 100, 2) + '%';
                    } else {
                        s += '<br/>' + this.series.name + ': ' +
                            formatNumber(this.y);
                    }
                });
                return s;
            },
            shared: true
        },
        plotOptions: {
            column: {
                grouping: false,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: '訓練目標人數',
            data: data1,
            pointPadding: 0.1
        }, {
            name: '開訓人數',
            data: data2,
            pointPadding: 0.2
        }, {
            name: '結訓人數',
            data: data3,
            pointPadding: 0.3
        }, {
            name: '達成率',
            type: 'spline',
            yAxis: 1,
            data: data4
        }
            /*, {
            name: '錄取率',
            type: 'spline',
            yAxis: 1,
            data: data5
        }*/
        ]
    });
}

//自辦在職：指標2_各分署辦理訓練班次統計
function creatChartB_2(containerID, data1, data2, data3, data4, data5) {
    $('#' + containerID).highcharts({
        chart: {
            type: 'column'
        },
        title: {
            text: '各分署辦理訓練班次統計'
        },
        xAxis: {
            categories: data5
        },
        yAxis: [{ // Primary yAxis
            labels: {
                formatter: function () {
                    return this.axis.defaultLabelFormatter.call(this);
                }
            },
            title: {
                text: '班級數'
            }
        }, { // Secondary yAxis
            title: {
                text: '百分比'
            },
            opposite: true,
            labels: {
                formatter: function () {
                    return formatFloat(this.value * 100, 2) + '%';
                }
            },
            min: 0,
            max: 1
        }],
        legend: {
            shadow: false
        },
        tooltip: {
            formatter: function () {
                var s = '<b>' + this.x + '</b>';
                $.each(this.points, function () {
                    if (this.series.name.indexOf('率') > 0) {
                        s += '<br/>' + this.series.name + ': ' +
                            formatFloat(this.y * 100, 2) + '%';
                    } else {
                        s += '<br/>' + this.series.name + ': ' +
                            formatNumber(this.y);
                    }
                });
                return s;
            },
            shared: true
        },
        plotOptions: {
            column: {
                grouping: false,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: '核定開訓班數',
            data: data1,
            pointPadding: 0.1
        }, {
            name: '已開訓班數',
            data: data2,
            pointPadding: 0.2
        }, {
            name: '訓練班數達成率',
            type: 'spline',
            yAxis: 1,
            data: data3
        }
        ]
    });
}

// Data gathered from http://populationpyramid.net/germany/2015/

//Age categories
var categories = [
    '14歲以下', '15~19歲', '20~24歲', '25~29歲',
    '30~34歲', '35~39歲', '40~44歲', '45~49歲',
    '50~54歲', '55~59歲', '60~64歲', '65歲以上'
];

function createMFChartB_3(containerID, data1, data2) {

    $('#' + containerID).highcharts({
        chart: {
            type: 'bar'
        },
        title: {
            text: '參訓年齡/性別分佈'
        },
        xAxis: [{
            categories: categories,
            reversed: false,
            labels: {
                step: 1
            }
        }, { // mirror axis on right side
            opposite: true,
            reversed: false,
            categories: categories,
            linkedTo: 0,
            labels: {
                step: 1
            }
        }],
        yAxis: {
            title: {
                text: null
            },
            labels: {
                formatter: function () {
                    //return Math.abs(this.value) + '%';
                    return Math.abs(this.value);
                }
            }
        },

        plotOptions: {
            series: {
                stacking: 'normal'
            }
        },

        tooltip: {
            formatter: function () {
                return '<b>' + this.series.name + ', ' + this.point.category + '</b><br/>' +
                    //'百分比: ' + formatFloat(Math.abs(this.point.y), 2) + '%';
                '人數: ' + formatFloat(Math.abs(this.point.y), 2);

            }
        },

        series: [{
            name: '男',
            data: data1
        }, {
            name: '女',
            data: data2
        }]
    });
}
//createMFChartB_3('VisualChartB_3',
//    [10, 20, 30, 50, 90, 70, 60, 33, 26, 10, 6, 2],
//    [8, 26, 22, 96, 102, 156, 88, 64, 24, 15, 3, 0],
//    vdata3, vdata4, vdata5);

/*
createMFChart('container2',
			[								//男 14歲以下~65歲以上
				-2.2, -2.1, -2.2, -2.4,
				-2.7, -3.0, -3.3, -3.2,
				-2.9, -3.5, -4.4, -4.1
			],[								//女 14歲以下~65歲以上
				2.1, 2.0, 2.1, 2.3, 
				2.9, 3.2, 3.1, 2.9, 
				4.3, 4.0, 3.5, 2.9
			]
);*/
//****************************************************************************************************//

//產投：指標1_總體參訓人數指標
function creatChartA_1(containerID, data1, data2, data3, data4, data5, data6, data7, data8) {
    $('#' + containerID).highcharts({
        chart: {
            type: 'column'
        },
        title: {
            text: '總體參訓人數指標'
        },
        xAxis: {
            categories: data1
        },
        yAxis: [{ // Primary yAxis
            labels: {
                formatter: function () {
                    return this.axis.defaultLabelFormatter.call(this);
                }
            },
            title: {
                text: '人數'
            }
        }, { // Secondary yAxis
            title: {
                text: '百分比'
            },
            opposite: true,
            labels: {
                formatter: function () {
                    return formatFloat(this.value * 100, 2) + '%';
                }
            },
            min: 0,
            max: 1
        }],
        legend: {
            shadow: false
        },
        tooltip: {
            formatter: function () {
                var s = '<b>' + this.x + '</b>';
                $.each(this.points, function () {
                    if (this.series.name.indexOf('率') > 0) {
                        s += '<br/>' + this.series.name + ': ' +
                            formatFloat(this.y * 100, 2) + '%';
                    } else {
                        s += '<br/>' + this.series.name + ': ' +
                            formatNumber(this.y);
                    }
                });
                return s;
            },
            shared: true
        },
        plotOptions: {
            column: {
                grouping: false,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: '核定人數',
            data: data2,
            pointPadding: 0.1
        }, {
            name: '報名人數',
            data: data3,
            pointPadding: 0.2
        }, {
            name: '開訓人數',
            data: data4,
            pointPadding: 0.3
        }, {
            name: '結訓人數',
            data: data5,
            pointPadding: 0.4
        }, {
            name: '撥款人數',
            data: data6,
            pointPadding: 0.5
        }, {
            name: '錄取率',
            type: 'spline',
            yAxis: 1,
            data: data7
        }, {
            name: '結訓率',
            type: 'spline',
            yAxis: 1,
            data: data8
        }
        ]
    });
}

//產投：指標2_總體開班數指標
function creatChartA_2(containerID, data1, data2, data3, data4, data5, data6, data7) {
    $('#' + containerID).highcharts({
        chart: {
            type: 'column'
        },
        title: {
            text: '總體開班數指標'
        },
        xAxis: {
            categories: data1
        },
        yAxis: [{ // Primary yAxis
            labels: {
                formatter: function () {
                    return this.axis.defaultLabelFormatter.call(this);
                }
            },
            title: {
                text: '班級數'
            }
        }, { // Secondary yAxis
            title: {
                text: '百分比'
            },
            opposite: true,
            labels: {
                formatter: function () {
                    return formatFloat(this.value * 100, 2) + '%';
                }
            },
            min: 0,
            max: 1
        }],
        legend: {
            shadow: false
        },
        tooltip: {
            formatter: function () {
                var s = '<b>' + this.x + '</b>';
                $.each(this.points, function () {
                    if (this.series.name.indexOf('率') > 0) {
                        s += '<br/>' + this.series.name + ': ' +
                            formatFloat(this.y * 100, 2) + '%';
                    } else {
                        s += '<br/>' + this.series.name + ': ' +
                            formatNumber(this.y);
                    }
                });
                return s;
            },
            shared: true
        },
        plotOptions: {
            column: {
                grouping: false,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: '提案班數',
            data: data2,
            pointPadding: 0.1
        }, {
            name: '核定班數',
            data: data3,
            pointPadding: 0.2
        }, {
            name: '開訓班數',
            data: data4,
            pointPadding: 0.3
        }, {
            name: '結訓班數',
            data: data5,
            pointPadding: 0.4
        }, {
            name: '開訓率',
            type: 'spline',
            yAxis: 1,
            data: data6
        }, {
            name: '結訓率',
            type: 'spline',
            yAxis: 1,
            data: data7
        }
        ]
    });
}

//產投：指標3_總體補助費指標
function creatChartA_3(containerID, data1, data2, data3, data4, data5, data6) {
    $('#' + containerID).highcharts({
        chart: {
            type: 'column'
        },
        title: {
            text: '總體補助費指標'
        },
        xAxis: {
            categories: data1
        },
        yAxis: [{ // Primary yAxis
            labels: {
                formatter: function () {
                    return this.axis.defaultLabelFormatter.call(this);
                }
            },
            title: {
                text: '金額'
            }
        }, { // Secondary yAxis
            title: {
                text: '百分比'
            },
            opposite: true,
            labels: {
                formatter: function () {
                    return formatFloat(this.value * 100, 2) + '%';
                }
            },
            min: 0,
            max: 1
        }],
        legend: {
            shadow: false
        },
        tooltip: {
            formatter: function () {
                var s = '<b>' + this.x + '</b>';
                $.each(this.points, function () {
                    if (this.series.name.indexOf('率') > 0) {
                        s += '<br/>' + this.series.name + ': ' +
                            formatFloat(this.y * 100, 2) + '%';
                    } else {
                        s += '<br/>' + this.series.name + ': ' +
                            formatNumber(this.y);
                    }
                });
                return s;
            },
            shared: true
        },
        plotOptions: {
            column: {
                grouping: false,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: '申請補助費',
            data: data2,
            pointPadding: 0.1
        }, {
            name: '核定補助費',
            data: data3,
            pointPadding: 0.2
        }, {
            name: '預估補助費',
            data: data4,
            pointPadding: 0.3
        }, {
            name: '核撥補助費',
            data: data5,
            pointPadding: 0.4
        }, {
            name: '執行率',
            type: 'spline',
            yAxis: 1,
            data: data6
        }
        ]
    });
}

//產投：指標4_19大類指標統計
function creatChartA_4(containerID, data1, data2, data3, data4, data5, data6, data7, data8) {
    $('#' + containerID).highcharts({
        chart: {
            type: 'column'
        },
        title: {
            text: '19大類指標統計'
        },
        xAxis: {
            categories: data1
        },
        yAxis: [{ // Primary yAxis
            labels: {
                formatter: function () {
                    return this.axis.defaultLabelFormatter.call(this);
                }
            },
            title: {
                text: '人數'
            }
        }, { // Secondary yAxis
            title: {
                text: '百分比'
            },
            opposite: true,
            labels: {
                formatter: function () {
                    return formatFloat(this.value * 100, 2) + '%';
                }
            },
            min: 0,
            max: 1
        }],
        legend: {
            shadow: false
        },
        tooltip: {
            formatter: function () {
                var s = '<b>' + this.x + '</b>';
                $.each(this.points, function () {
                    if (this.series.name.indexOf('率') > 0) {
                        s += '<br/>' + this.series.name + ': ' +
                            formatFloat(this.y * 100, 2) + '%';
                    } else {
                        s += '<br/>' + this.series.name + ': ' +
                            formatNumber(this.y);
                    }
                });
                return s;
            },
            shared: true
        },
        plotOptions: {
            column: {
                grouping: false,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: '核定人數',
            data: data2,
            pointPadding: 0.1
        }, {
            name: '報名人數',
            data: data3,
            pointPadding: 0.2
        }, {
            name: '開訓人數',
            data: data4,
            pointPadding: 0.3
        }, {
            name: '結訓人數',
            data: data5,
            pointPadding: 0.4
        }, {
            name: '撥款人數',
            data: data6,
            pointPadding: 0.5
        }, {
            name: '錄取率',
            type: 'spline',
            yAxis: 1,
            data: data7
        }, {
            name: '結訓率',
            type: 'spline',
            yAxis: 1,
            data: data8
        }
        ]
    });
}

//產投：指標5_政策性產業參訓人數統計
function creatChartA_5(containerID, data1, data2, data3, data4, data5, data6, data7, data8) {
    $('#' + containerID).highcharts({
        chart: {
            type: 'column'
        },
        title: {
            text: '政策性產業參訓人數統計'
        },
        xAxis: {
            categories: data1
        },
        yAxis: [{ // Primary yAxis
            labels: {
                formatter: function () {
                    return this.axis.defaultLabelFormatter.call(this);
                }
            },
            title: {
                text: '人數'
            }
        }, { // Secondary yAxis
            title: {
                text: '百分比'
            },
            opposite: true,
            labels: {
                formatter: function () {
                    return formatFloat(this.value * 100, 2) + '%';
                }
            },
            min: 0,
            max: 1
        }],
        legend: {
            shadow: false
        },
        tooltip: {
            formatter: function () {
                var s = '<b>' + this.x + '</b>';
                $.each(this.points, function () {
                    if (this.series.name.indexOf('率') > 0) {
                        s += '<br/>' + this.series.name + ': ' +
                            formatFloat(this.y * 100, 2) + '%';
                    } else {
                        s += '<br/>' + this.series.name + ': ' +
                            formatNumber(this.y);
                    }
                });
                return s;
            },
            shared: true
        },
        plotOptions: {
            column: {
                grouping: false,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: '核定人數',
            data: data2,
            pointPadding: 0.1
        }, {
            name: '報名人數',
            data: data3,
            pointPadding: 0.2
        }, {
            name: '開訓人數',
            data: data4,
            pointPadding: 0.3
        }, {
            name: '結訓人數',
            data: data5,
            pointPadding: 0.4
        }, {
            name: '撥款人數',
            data: data6,
            pointPadding: 0.5
        }, {
            name: '錄取率',
            type: 'spline',
            yAxis: 1,
            data: data7
        }, {
            name: '結訓率',
            type: 'spline',
            yAxis: 1,
            data: data8
        }
        ]
    });
}

//產投：指標6_政策性產業班級統計
function creatChartA_6(containerID, data1, data2, data3, data4, data5, data6, data7) {
    $('#' + containerID).highcharts({
        chart: {
            type: 'column'
        },
        title: {
            text: '政策性產業班級統計'
        },
        xAxis: {
            categories: data1
        },
        yAxis: [{ // Primary yAxis
            labels: {
                formatter: function () {
                    return this.axis.defaultLabelFormatter.call(this);
                }
            },
            title: {
                text: '班級數'
            }
        }, { // Secondary yAxis
            title: {
                text: '百分比'
            },
            opposite: true,
            labels: {
                formatter: function () {
                    return formatFloat(this.value * 100, 2) + '%';
                }
            },
            min: 0,
            max: 1
        }],
        legend: {
            shadow: false
        },
        tooltip: {
            formatter: function () {
                var s = '<b>' + this.x + '</b>';
                $.each(this.points, function () {
                    if (this.series.name.indexOf('率') > 0) {
                        s += '<br/>' + this.series.name + ': ' +
                            formatFloat(this.y * 100, 2) + '%';
                    } else {
                        s += '<br/>' + this.series.name + ': ' +
                            formatNumber(this.y);
                    }
                });
                return s;
            },
            shared: true
        },
        plotOptions: {
            column: {
                grouping: false,
                shadow: false,
                borderWidth: 0
            }
        },
        series: [{
            name: '提案班數',
            data: data2,
            pointPadding: 0.1
        }, {
            name: '核定班數',
            data: data3,
            pointPadding: 0.2
        }, {
            name: '開訓班數',
            data: data4,
            pointPadding: 0.3
        }, {
            name: '結訓班數',
            data: data5,
            pointPadding: 0.4
        }, {
            name: '開訓率',
            type: 'spline',
            yAxis: 1,
            data: data6
        }, {
            name: '結訓率',
            type: 'spline',
            yAxis: 1,
            data: data7
        }
        ]
    });
}

//產投：指標9_三年補助使用情形
function creatChartA_9(containerID, SubCategories, data2, data3, data4) {
    var data = [data2, data3, data4];
    var colors = Highcharts.getOptions().colors,
        categories = ["第一次使用", "第二~五次使用", "第六次(含)以上使用"],
        SubCategories,
        browserData = [],
        versionsData = [],
        i,
        j,
        dataLen = data.length,
        drillDataLen,
        brightness;

    // Build the data arrays
    for (i = 0; i < dataLen; i += 1) {
        total = 0;
        // add version data
        drillDataLen = SubCategories.length;//data[i].drilldown.data.length;
        for (j = 0; j < drillDataLen; j += 1) {
            brightness = 0.2 - (j / drillDataLen) / 5;
            versionsData.push({
                name: SubCategories[j],
                y: data[i][j],
                color: Highcharts.Color(colors[i]).brighten(brightness).get()
            });
            total += data[i][j];
        }

        // add browser data
        browserData.push({
            name: categories[i],
            y: total,
            color: colors[i]
        });
    }

    // Create the chart
    $('#' + containerID).highcharts({
        chart: {
            type: 'pie'
        },
        title: {
            text: '三年補助使用情形'
        },
        yAxis: {
            title: {
                text: 'Total percent market share'
            }
        },
        plotOptions: {
            pie: {
                shadow: false,
                center: ['50%', '50%']
            }
        },
        tooltip: {
            valueSuffix: '%'
        },
        series: [{
            name: '百分比',
            data: browserData,
            size: '60%',
            dataLabels: {
                formatter: function () {
                    return this.y > 5 ? this.point.name : null;
                },
                color: '#ffffff',
                distance: -30
            }
        }, {
            name: '百分比',
            data: versionsData,
            size: '80%',
            innerSize: '60%',
            dataLabels: {
                formatter: function () {
                    // display only if larger than 1
                    return this.y > 1 ? '<b>' + this.point.name + ':</b> ' +
                        this.y + '%' : null;
                }
            },
            id: 'versions'
        }],
        responsive: {
            rules: [{
                condition: {
                    maxWidth: 400
                },
                chartOptions: {
                    series: [{
                        id: 'versions',
                        dataLabels: {
                            enabled: false
                        }
                    }]
                }
            }]
        }
    });

}


function creatDountChartA_9(containerID, data) {
    var colors = Highcharts.getOptions().colors,
        categories = [
            "第一次使用",
            "第二~五次使用",
            "第六次(含)以上使用"
        ],
        SubCategories = [
            "北基宜花金馬分署",
            "桃竹苗分署",
            "中彰投分署",
            "雲嘉南分署",
            "高屏澎東分署"
        ],
    browserData = [],
    versionsData = [],
    i,
    j,
    dataLen = data.length,
    drillDataLen,
    brightness;

    // Build the data arrays
    for (i = 0; i < dataLen; i += 1) {
        total = 0;
        // add version data
        drillDataLen = SubCategories.length;//data[i].drilldown.data.length;
        for (j = 0; j < drillDataLen; j += 1) {
            brightness = 0.2 - (j / drillDataLen) / 5;
            versionsData.push({
                name: SubCategories[j],
                y: data[i][j],
                color: Highcharts.Color(colors[i]).brighten(brightness).get()
            });
            total += data[i][j];
        }

        // add browser data
        browserData.push({
            name: categories[i],
            y: total,
            color: colors[i]
        });
    }

    // Create the chart
    $('#' + containerID).highcharts({
        chart: {
            type: 'pie'
        },
        title: {
            text: '三年補助使用情形'
        },
        yAxis: {
            title: {
                text: 'Total percent market share'
            }
        },
        plotOptions: {
            pie: {
                shadow: false,
                center: ['50%', '50%']
            }
        },
        tooltip: {
            valueSuffix: '%'
        },
        series: [{
            name: '百分比',
            data: browserData,
            size: '60%',
            dataLabels: {
                formatter: function () {
                    return this.y > 5 ? this.point.name : null;
                },
                color: '#ffffff',
                distance: -30
            }
        }, {
            name: '百分比',
            data: versionsData,
            size: '80%',
            innerSize: '60%',
            dataLabels: {
                formatter: function () {
                    // display only if larger than 1
                    return this.y > 1 ? '<b>' + this.point.name + ':</b> ' +
                        this.y + '%' : null;
                }
            },
            id: 'versions'
        }],
        responsive: {
            rules: [{
                condition: {
                    maxWidth: 400
                },
                chartOptions: {
                    series: [{
                        id: 'versions',
                        dataLabels: {
                            enabled: false
                        }
                    }]
                }
            }]
        }
    });

}
/*
creatDountChart('container3',
			[[0.1, 1.3, 53.02, 1.4, 0.88], //第一次使用 			北基~高屏 百分比
    		[1.02, 7.36, 0.35, 0.11, 0.1 ],	//第二~五次使用			北基~高屏 百分比
            [6.2, 0.29, 0.27,0.47, 0.47]]	//第六次(含)以上使用	北基~高屏 百分比
);*/
//****************************************************************************************************//





function creatStackChart(containerID, data1, data2, data3, data4) {
    $('#' + containerID).highcharts({

        chart: {
            type: 'column'
        },

        title: {
            text: '各分署三年補助使用情形'
        },

        xAxis: {
            categories: ['北基宜花金馬分署', '桃竹苗分署', '中彰投分署', '雲嘉南分署', '高屏澎東分署']
        },

        yAxis: {
            allowDecimals: false,
            min: 0,
            labels: {
                formatter: function () {
                    return this.axis.defaultLabelFormatter.call(this);
                }
            },
            title: {
                text: '人數'
            }
        },

        tooltip: {
            formatter: function () {
                return '<b>' + this.x + '</b><br/>' +
                    this.series.name + ': ' + this.y + '<br/>' +
                    'Total: ' + this.point.stackTotal;
            }
        },

        plotOptions: {
            column: {
                stacking: 'normal'
            }
        },

        series: [{
            name: '三年內使用經費已達八成以上的學員數',
            data: data1,
            stack: '三年內'
        }, {
            name: '三年內使用經費未達八成的學員數',
            data: data2,
            stack: '三年內'
        }, {
            name: '一年內使用經費已達八成以上的學員數',
            data: data3,
            stack: '一年內'
        }, {
            name: '一年內使用經費未達八成的學員數',
            data: data4,
            stack: '一年內'
        }]
    });
}
/*creatStackChart('container4',
[5, 3, 4, 7, 2],	//三年內使用經費已達八成以上的學員數
[3, 4, 4, 2, 5],	//三年內使用經費未達八成的學員數
[2, 5, 6, 2, 1],	//一年內使用經費已達八成以上的學員數
[3, 0, 4, 4, 3]		//一年內使用經費未達八成的學員數
);*/
//****************************************************************************************************//









