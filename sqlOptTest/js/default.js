let xmlHttp = null;
const urlHead = "/service.ashx";

//获取数据
function getData() {
    xmlHttp = new XMLHttpRequest();
    let mobile = document.getElementById("qrymobile").value;
    let url = urlHead + "?optType=qry&&mobile=" + mobile;
    xmlHttp.open("GET", url, false);
    xmlHttp.send(null);
    document.getElementById("log").innerHTML = xmlHttp.responseText;
}

function getDatap() {
    xmlHttp = new XMLHttpRequest();
    let url = urlHead;
    let mobile = document.getElementById("qrymobile").value;
    xmlHttp.open("post", url, false);
    xmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=UTF-8");
    xmlHttp.send("optType=qry&mobile=" + mobile);
    document.getElementById("log").innerHTML = xmlHttp.responseText;
}

//增加数据
function addData() {
    xmlHttp = new XMLHttpRequest();
    let account = document.getElementById("account").value;
    let mobile = document.getElementById("mobile").value;
    let parms = "?optType=ins&&mobile=" + mobile + "&&account=" + account;
    let url = urlHead + parms;
    xmlHttp.open("GET", url, false);
    xmlHttp.send(null);
    document.getElementById("log").innerHTML = xmlHttp.responseText;
}

function addDatap() {
    xmlHttp = new XMLHttpRequest();
    let account = document.getElementById("account").value;
    let mobile = document.getElementById("mobile").value;
    let url = urlHead;
    xmlHttp.open("POST", url, false);
    xmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=UTF-8");
    xmlHttp.send("optType=ins&mobile=" + mobile + "&account=" + account);
    document.getElementById("log").innerHTML = xmlHttp.responseText;
}

//编辑数据
function edtData() {

    xmlHttp = new XMLHttpRequest();
    let account = document.getElementById("edtaccount").value;
    let mobile = document.getElementById("edtmobile").value;
    let url = urlHead + "?optType=edt&&mobile=" + mobile + "&&account=" + account;
    xmlHttp.open("GET", url, false);
    xmlHttp.send(null);
    document.getElementById("log").innerHTML = xmlHttp.responseText;
}
function edtDatap() {

    xmlHttp = new XMLHttpRequest();
    let account = document.getElementById("edtaccount").value;
    let mobile = document.getElementById("edtmobile").value;
    let url = urlHead;
    xmlHttp.open("POST", url, false);
    xmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=UTF-8");
    xmlHttp.send("optType=edt &mobile=" + mobile + "&account=" + account);
    document.getElementById("log").innerHTML = xmlHttp.responseText;
}

//删除数据
function delData() {
    xmlHttp = new XMLHttpRequest();
    let url = urlHead + "?optType=del";
    xmlHttp.open("GET", url, false);
    xmlHttp.send(null);
    if (xmlHttp.responseText > 0) {
        document.getElementById("log").innerHTML = '成功删除' + xmlHttp.responseText + "条数据";
    } else {
        document.getElementById("log").innerHTML = '数据删除失败';
    }
}
function delDatap() {
    xmlHttp = new XMLHttpRequest();
    let url = urlHead;
    xmlHttp.open("POST", url, false);
    xmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=UTF-8");

    xmlHttp.send("optType=del");
    if (xmlHttp.responseText > 0) {
        document.getElementById("log").innerHTML = '成功删除' + xmlHttp.responseText + "条数据";
    } else {
        document.getElementById("log").innerHTML = '数据删除失败';
    }
}

//执行存储过程
function proc() {
    xmlHttp = new XMLHttpRequest();
    let url = urlHead + "?procName=qry_sys_user&&id=0001&&optType=proc";
    xmlHttp.open("GET", url, false);
    xmlHttp.send(null);
    document.getElementById("log").innerHTML = xmlHttp.responseText;
}
function procp() {
    xmlHttp = new XMLHttpRequest();
    let url = urlHead;
    xmlHttp.open("POST", url, false);
    xmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=UTF-8");
    xmlHttp.send("procName=qry_sys_user&&id=0001&&optType=proc");
    document.getElementById("log").innerHTML = xmlHttp.responseText;
}

function imData() {
    var path = document.getElementById("imData").value;
    if ($.trim(path) == "") { alert("请选择要上传的文件"); return; }
    console.log("test");
    debugger
    var result_msg = "";
    $.ajax({
        url: 'files/impData.ashx',  //这里是服务器处理的代码
        type: 'post',
        secureuri: false, //一般设置为false
        fileElementId: 'fu_UploadFile', // 上传文件的id、name属性名
        dataType: 'text', //返回值类型，一般设置为json、application/json
        data: { path: path, name: "" }, //传递参数到服务器
        success: function (data, status) {
            console.log(data)
        },
        error: function (data, status, e) {
            console.log(data)
            alert("错误：上传组件错误，请检察网络!");
        }
    });
}
