//for PIE
var tipRenderer = function (storeItem, item) {
    //calculate percentage.
    var total = 0;

    App.Chart1.getStore().each(function (rec) {
        total += rec.get('Data1');
    });

    this.setTitle(storeItem.get('Name') + ': ' + Math.round(storeItem.get('Data1') / total * 100) + '%');
};

function getStatus(record) {
    var sReturn = "";
    var score = record.get("score");
    var p = score / 100;
    if (p >= 0.8) {
        sReturn = "Excellent(" + score + ")";
    }
    else if (p >= 0.6) {
        sReturn = "Good(" + score + ")";
    }
    else {
        sReturn = "Poor(" + score + ")";
    }
    //added 2013-09-06 
    var myDate = new Date();
    if (myDate.getDate() > 0 && myDate.getDate() <11) {
        sReturn = "Evaluating";
    }
    return rowFormat(sReturn,record);
}
function selectEvent(record, StoreSub) {
    //StoreSub.loadRecords(record);            
    //StoreSub.reload();
    App.hiddenSelectedNO.setValue(record.get("NO"));
    //App.fileContractNO.setValue(record.get("NO"));//App.RowSelectionModelMain.getSelection()[0].data.NO
    mydata = [{ 'workCenter': record.get('workCenter'), 'contractor': record.get('contractor'), 'contractAdmin': record.get('contractAdmin'), 'buyer': record.get('buyer'), 'mainCoordinator': record.get('mainCoordinator'), 'userRepresentative': record.get('userRepresentative'), 'contact': record.get('contact'), 'telephone': record.get('telephone'), 'validateDate': record.get('validateDate'), 'expireDate': record.get('expireDate') }];
    StoreSub.loadData(mydata, false);
    
    App.storePie1.reload({
        params: { NO: record.get("NO") }
    });
    App.storePie2.reload({
        params: { NO: record.get("NO") }
    });
    App.storeEvaluation.reload({
        params: { NO: record.get("NO") }
    });
    App.storeFiles.reload({
        params: { NO: record.get("NO")}
    });
    App.storeKPI2Scatter.reload({
        params: { NO: record.get("NO"), DateStart: App.dfKPI2Start.getValue(), DateEnd: App.dfKPI2End.getValue() }
    });
    App.storeKPI3.reload({
        params: { NO: record.get("NO"), DateStart: App.dfKPI3Start.getValue(), DateEnd: App.dfKPI3End.getValue() }
    });
    App.pContractor.body.update("<br/><center>" + record.get("contractor") + "</center>");
}
function moneyFormat(value, senderObject, record, rowIndex, columnIndex) {
    var sReturn = value;
    if (columnIndex == 5 || columnIndex == 6 || columnIndex == 7) {
        value = value.toFixed(0);
        n = String(value);
        re = /(\d{1,3})(?=(\d{3})+(?:$|\.))/g;
        n1 = n.replace(re, "$1,");
        sReturn= n1;
    }
    sReturn = rowFormat(sReturn,record);
    return sReturn;
}
function moneyFormat2(value) {
    try {
    var sReturn = value;
    value = value.toFixed(0);
    n = String(value);
    re = /(\d{1,3})(?=(\d{3})+(?:$|\.))/g;
    n1 = n.replace(re, "$1,");
    sReturn = n1;
    return sReturn;
    } catch (e) {
        return value;
    }
}
function percentageFormat(value, senderObject, record, rowIndex, columnIndex) {
    var tmpValue = value * 100;
    tmpValue = tmpValue.toFixed(2);
    if (value < 0.1 && record.get("expireDate") > (new Date())) {
        return '<font color="red">' + tmpValue + '%</font>';
    }
    return rowFormat(tmpValue + "%",record);
}
function rowFormat(value,record)
{
    var sReturn = "";
    var dateInterval=record.get("expireDate") - (new Date());
    var days = Math.floor(dateInterval / (24 * 3600 * 1000))
    if (days<0) {
        sReturn = "gray";
    }
    else if (days < 30) {
        sReturn = "maroon";
    }
    if (sReturn == "") {
        return value;
    }
    sReturn='<font color="'+sReturn+'">' + value + '</font>';
    return sReturn;
}
function searchClick(realSearching)
{
    //var sNO = App.searchNO.getValue();
    //var sCA = App.searchCA.getValue();
    //var sSPS = App.SearchPringScheme.getValue();
    
    //App.storeMain.load({
    //    params: { NO: sNO, contractAdmin: sCA, pricingScheme: sSPS }
    //});
    App.winSearch.hide();
    App.hiddenValue.setValue("1");
    App.storeMain.load();
    if (realSearching==false) {
        clearSearch();
    }
}
function preLoad()
{    
    var sIsSearching = App.hiddenValue.getValue();
    var sKeyword,sNO, sCA, sSPS;
    if (sIsSearching == "1") {
        sKeyword = App.searchKeyword.getValue();
        sNO = App.searchNO.getValue();
        sCA = App.searchCA.getValue();
        sSPS = App.SearchPringScheme.getValue();
        App.cbValid.setValue(false);
    }
    var isValidOnly = App.cbValid.getValue();
    var new_params = { searchKey:sKeyword,searchNO: sNO, CA: sCA, PS: sSPS, isValid: isValidOnly };
    Ext.apply(App.storeMain.proxy.extraParams, new_params);
}
function afterLoad()
{
    if (App.storeMain.getCount() == 0) { showInformation('Information', 'No results found!') };
}
function clearSearch()
{
    App.searchKeyword.clear();
    App.searchNO.clear();
    App.searchCA.clear();
    App.SearchPringScheme.clear();
    App.hiddenValue.setValue("0");
}

function hilightEvaluation()
{
    App.tp1.setActiveTab(1);
    updateSpot(App.gridEvaluationContract);
}
var updateSpot = function (cmp) {
    App.Spot.show(cmp);
    App.Spot.hide();
    if (App.Spot.active) {
        App.Spot.hide();
    }
};

function updateEvaluation(record)
{
    if (record.dirty) {
        //var data = App.storeEvaluationContract.getChangedData();

    }
}

function searchAutomatic(Contract_No)
{
    clearSearch();
    App.searchNO.setValue(Contract_No);    
    searchClick(false);
}

function showMsg(sTitle,sCon)
{
    Ext.net.Notification.show({
        iconCls: '#Information',
        pinEvent: 'click',
        html: sCon,
        title: sTitle
    });
}
function showAlert(sTitle,sCon)
{
    Ext.Msg.show({"title":sTitle,"buttons":Ext.Msg.OK,"icon":Ext.Msg.ERROR,"msg":sCon});
}
function showInformation(sTitle, sCon) {
    Ext.Msg.show({ "title": sTitle, "buttons": Ext.Msg.OK, "icon": Ext.Msg.INFO, "msg": sCon });
}

function evaluateRender(value,senderObj,record)
{
    if (record.data.Evaluated) {
        return '<font color="gray">' + record.data.Contract_No + '</font>';
    }
    return record.data.Contract_No;
}
var onShow = function (toolTip, grid) {
    var view = grid.getView(),
        record = view.getRecord(toolTip.triggerElement),
        //data = Ext.encode(record.data);
        data = "File Name:" + record.data.FileName + "<br/>Remark:"+record.data.Remark+"<br/>双击下载!<br/>Double Click to download!";
    toolTip.update(data);
};
var onShowMain = function (toolTip, grid) {
    if (App.rdUser.getValue()) {
        toolTip.setVisible(true);
    } else {
        toolTip.setVisible(false); return;
    }
    var remark = "零星为不可容忍，一星很差，二星差，三星一般，四星好，五星优秀<br/>当评价结果为零星，一星或五星时，用户必须上载附件以说明具体原因";
    var view = grid.getView(),
        store = grid.getStore(),
        record = view.getRecord(view.findItemByChild(toolTip.triggerElement)),
        column = view.getHeaderByCell(toolTip.triggerElement),
        data = record.get(column.dataIndex);//indexOf Evaluated

    if (column.text.indexOf("Evaluate") > 0) { toolTip.update("点击提交测评<br/>Click to Evaluate"); return; };
    if (column.text.indexOf("Upload Files") > 0) { toolTip.update("上传文件(如果存在)<br/>Upload files if any"); return; };
    if (column.dataIndex.indexOf("Score") < 0) { toolTip.update(data); return; }
    //toolTip.enable(); toolTip.show();
    if (record.data.Evaluated) {
        switch (data)
        {
            case 0:
                remark = "不可容忍<br/>请上传相关附件";
                break;
            case 1:
                remark = "很差<br/>请上传相关附件";
                break;
            case 2:
                remark = "差";
                break;
            case 3:
                remark = "一般";
                break;
            case 4:
                remark = "好";
                break;
            case 5:
                remark = "优秀<br/>请上传相关附件";
                break;
            default:
                break;
        }
        toolTip.update(remark); return;
    }
    else {
        if (data == 0) { toolTip.update(remark); return;}
    }
    //return;
    //toolTip.update(column.dataIndex);
    //toolTip.hide();
};
function loginTypeChange()
{
    //App.storeEvaluationContract.removeAll();
    if(App.rdUser.getValue())
    {
        App.btnDownloadTemplate.disable(); App.btnBatchImport.disable();
        App.RatingColumn1.show(); App.RatingColumn2.show(); App.RatingColumn3.show();
        App.RatingColumn4.show(); App.RatingColumn5.show(); App.RatingColumn6.show();
        App.clmEvaluate.show(); App.clmDepScore.hide();
    }
    else
    {
        App.btnDownloadTemplate.enable(); App.btnBatchImport.enable();
        App.RatingColumn1.hide(); App.RatingColumn2.hide(); App.RatingColumn3.hide();
        App.RatingColumn4.hide(); App.RatingColumn5.hide(); App.RatingColumn6.hide();
        App.clmEvaluate.hide();App.clmDepScore.show();
    }
}
function checkLogin(msg)
{
    var sUser = App.hiddenUser.getValue();
    if (sUser != null && sUser != "") {
        return true;
    }
    else {
        if (msg==null) {
            showMsg("Information", "You are not logged in!");
        }
        return false;
    }
}

var saveChart = function (btn) {
    Ext.MessageBox.confirm('Confirm Download', 'Would you like to download the chart as an image?', function (choice) {
        if (choice == 'yes') {
            btn.up('panel').down('chart').save({
                type: 'image/png'
            });
        }
    });
}
var onKeyUp = function () {
    var me = this,
        v = me.getValue(),
        field;

    if (me.startDateField) {
        field = Ext.getCmp(me.startDateField);
        field.setMaxValue(v);
        me.dateRangeMax = v;
    } else if (me.endDateField) {
        field = Ext.getCmp(me.endDateField);
        field.setMinValue(v);
        me.dateRangeMin = v;
    }

    field.validate();
};
function submitTimeRange(btn)
{
    //showMsg("ID", btn.getId());App.DateField2.setValue(new Date(2013,7,9,9,10,17,775));new Date(yyyy,mth,dd);
    var dateNow = new Date();
    switch(btn.getId())
    {
        case "btnTimeRangeSubmit":
            
            break;
        case "MenuItem1":
            App.DateField1.setValue((new Date()).dateAdd('m', -2));
            App.DateField2.setValue(dateNow);
            break;
        case "MenuItem2":
            App.DateField1.setValue((new Date()).dateAdd('m', -5));
            App.DateField2.setValue(dateNow);
            break;
        case "MenuItem3":
            App.DateField1.setValue((new Date()).dateAdd('m', -11));
            App.DateField2.setValue(dateNow);
            break;
        default:
            
    }
    App.storePie2.reload();
}
function exportClick(btn)
{
    var exportType="";
    switch (btn.getId()) {
        case "btnEvaluationExport":
            exportType = "overall";
            break;
        case "MenuItem4":
            exportType = "month";
            break;
        case "MenuItem5":
            exportType = "quarter";
            break;
        case "MenuItem6":
            exportType = "year";
            break;
        case "MenuItem7":
            exportType = "overall";
            break;
        case "MenuItem8":
            exportType = "monthdata";
            break;
        default:

    }
    window.open('report.aspx?FO=' + App.hiddenSelectedNO.getValue() + '&startDate=' + App.DateField1.getValue().format('yyyy-MM-dd')+'&endDate='+App.DateField2.getValue().format('yyyy-MM-dd')+'&export='+exportType);
}

/* 得到日期年月日等加数字后的日期 */
Date.prototype.dateAdd = function (interval, number) {
    var d = this;
    var k = { 'y': 'FullYear', 'q': 'Month', 'm': 'Month', 'w': 'Date', 'd': 'Date', 'h': 'Hours', 'n': 'Minutes', 's': 'Seconds', 'ms': 'MilliSeconds' };
    var n = { 'q': 3, 'w': 7 };
    eval('d.set' + k[interval] + '(d.get' + k[interval] + '()+' + ((n[interval] || 1) * number) + ')');
    return d;
}
/* 计算两日期相差的日期年月日等 [#]BMsrA91ARVfg4agApe2PEA==[#]*/
Date.prototype.dateDiff = function (interval, objDate2) {
    var d = this, i = {}, t = d.getTime(), t2 = objDate2.getTime();
    i['y'] = objDate2.getFullYear() - d.getFullYear();
    i['q'] = i['y'] * 4 + Math.floor(objDate2.getMonth() / 4) - Math.floor(d.getMonth() / 4);
    i['m'] = i['y'] * 12 + objDate2.getMonth() - d.getMonth();
    i['ms'] = objDate2.getTime() - d.getTime();
    i['w'] = Math.floor((t2 + 345600000) / (604800000)) - Math.floor((t + 345600000) / (604800000));
    i['d'] = Math.floor(t2 / 86400000) - Math.floor(t / 86400000);
    i['h'] = Math.floor(t2 / 3600000) - Math.floor(t / 3600000);
    i['n'] = Math.floor(t2 / 60000) - Math.floor(t / 60000);
    i['s'] = Math.floor(t2 / 1000) - Math.floor(t / 1000);
    return i[interval];
}
Date.prototype.format = function (format) {
    var o = {
        "M+": this.getMonth() + 1, //month
        "d+": this.getDate(),    //day
        "h+": this.getHours(),   //hour
        "m+": this.getMinutes(), //minute
        "s+": this.getSeconds(), //second
        "q+": Math.floor((this.getMonth() + 3) / 3),  //quarter
        "S": this.getMilliseconds() //millisecond
    }
    if (/(y+)/.test(format)) format = format.replace(RegExp.$1,
    (this.getFullYear() + "").substr(4 - RegExp.$1.length));
    for (var k in o) if (new RegExp("(" + k + ")").test(format))
        format = format.replace(RegExp.$1,
        RegExp.$1.length == 1 ? o[k] :
        ("00" + o[k]).substr(("" + o[k]).length));
    return format;
}
var showResultText = function (btn, text) {
    if (btn == "ok") {
        App.direct.setKPIMOM(App.hiddenSelectedNO.getValue(), text, App.hiddenDateMonth.getValue());
    }
};
var showQuarterText = function () {
    var now = new Date;
    var nowMonth = now.getMonth();
    var nowYear = now.getFullYear();
    var sQuarter = "1st";
    if (nowMonth < 3) { sQuarter = "4th"; nowYear = nowYear-1 }
    else if (nowMonth < 6) { sQuarter = "1st"; }
    else if (nowMonth < 9) { sQuarter = "2nd"; }
    else if (nowMonth < 12) { sQuarter = "3rd"; }
    else { sQuarter = "4th";}
    return sQuarter + " Quarterly Report " + nowYear;
}
var scatterStoreReload = function (store, dateStart, dateEnd, isOverall) {
    store.removeAll();
    store.reload({ params: { DateStart: dateStart, DateEnd: dateEnd ,OverAll:isOverall} });
}

function test(obj, newValue, oldValue)
{
    //obj.disable();#{fileContractNO}.setValue(record.data.Contract_No);#{winFile}.show();#{storeWinFile}.reload();
    return;
}