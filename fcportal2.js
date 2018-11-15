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
    return rowFormat(sReturn,record);
}
function selectEvent(record, StoreSub) {
    //StoreSub.loadRecords(record);            
    //StoreSub.reload();
    mydata = [{ 'workCenter': record.get('workCenter'), 'contractor': record.get('contractor'), 'contractAdmin': record.get('contractAdmin'), 'buyer': record.get('buyer'), 'mainCoordinator': record.get('mainCoordinator'), 'userRepresentative': record.get('userRepresentative'), 'contact': record.get('contact'), 'telephone': record.get('telephone'), 'validateDate': record.get('validateDate'), 'expireDate': record.get('expireDate') }];
    StoreSub.loadData(mydata, false);
    
    App.storePie1.reload({
        params: { NO: record.get("NO") }
    });
    App.storePie2.reload({
        params: { NO: record.get("NO") }
    });
    App.storeEvaluation.reload({
        params: { NO: record.get("NO"), User: App.hiddenUser.value }
    });
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
function searchClick()
{
    var sNO = App.searchNO.getValue();
    var sCA = App.searchCA.getValue();
    var sSPS = App.SearchPringScheme.getValue();
    
    //App.storeMain.load({
    //    params: { NO: sNO, contractAdmin: sCA, pricingScheme: sSPS }
    //});
    App.winSearch.hide();
    App.hiddenValue.setValue("1");
    App.storeMain.load();
}
function preLoad()
{
    var isValidOnly = App.cbValid.getValue();
    var sIsSearching = App.hiddenValue.getValue();
    var sNO, sCA, sSPS;
    if (sIsSearching=="1") {
        sNO = App.searchNO.getValue();
        sCA = App.searchCA.getValue();
        sSPS = App.SearchPringScheme.getValue();
    }
    var new_params = { searchNO: sNO, CA: sCA, PS: sSPS, isValid: isValidOnly };
    Ext.apply(App.storeMain.proxy.extraParams, new_params);
}
function clearSearch()
{
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
    App.searchNO.setValue(Contract_No);
    searchClick();
}

function showMsg(sTitle,sCon)
{
    Ext.net.Notification.show({
        iconCls: 'icon-information',
        pinEvent: 'click',
        html: sCon,
        title: sTitle
    });
}

function test(obj,record)
{
    App.winFile.show();
    return;
}