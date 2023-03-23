//这个函数在整个wps加载项中是第一个执行的
function OnAddinLoad(ribbonUI) {
    wps.ribbonUI = ribbonUI
    if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
        wps.Enum = WPS_Enum
    }
    //往PluginStorage中设置一个标记，用于控制两个按钮的置灰
    wps.PluginStorage.setItem(constom_Enum.EnableFlag, false)
    //启动时常常显示
    // let tskpan_login = wps.CreateTaskPane("mail.163.com", "我是个登录框")
    // let id = tskpan_login.ID
    // wps.PluginStorage.setItem(constom_Enum.TaskPane_Login_ID, id)
    // tskpan_login.Visible = true
    // tskpan_login.Width = 600 * window.devicePixelRatio
    // tskpan_login.DockPosition = WPS_Enum.msoCTPDockPositionLeft

    //注册窗口激活时的事件
    wps.ApiEvent.AddApiEventListener("WindowActivate", OnWindowActivate)
    wps.ApiEvent.AddApiEventListener("ContentChange", OnContentChange)
    wps.ApiEvent.AddApiEventListener("DocumentBeforeCopy", OnDocumentBeforeCopy)
    wps.ApiEvent.AddApiEventListener("DocumentBeforePaste", OnDocumentBeforePaste)
    // wps.ApiEvent.AddApiEventListener("")
    //设置为修订状态
    // setTrackRevisions()

    return true
}
function OnDocumentBeforePaste(){
     console.log("OnDocumentBeforePaste")
}
function OnDocumentBeforeCopy(){
     console.log("OnDocumentBeforeCopy")
}
/**
 * 当文档内容发生变化时的事件
 *
 */
function OnContentChange() {
    console.log(wps.WpsApplication().ActiveDocument.Range(0, 20).Text)
}
/**
 * 窗口激活时事件实现
 *
 */
function OnWindowActivate() {
    console.log("激活的活动文档名称是：" + wps.WpsApplication().ActiveDocument.Name)
}
/**
 * 设置文档编辑模式为修订状态
 *
 */
function setTrackRevisions() {
    var l_doc = wps.WpsApplication().ActiveDocument;
    if (l_doc) {
        //设置文档为修订状态
        l_doc.TrackRevisions = true;
        let l_windowActivePaneView = wps.WpsApplication().ActiveWindow.ActivePane.View;
        //设置文档显示状态为「显示为最终标记状态」
        l_windowActivePaneView.RevisionsView = 0; //Wps.WpsWdRevisionsView.wdRevisionsViewFinal;
        //设置文档显示修订和批注
        l_windowActivePaneView.ShowRevisionsAndComments = true;
        //设置文档使用批注框显示标记
        l_windowActivePaneView.RevisionsMode = 0; //Wps.WpsWdRevisionsMode.wdBalloonRevisions;
    }
}
/**
 * 控件的动作
 *
 * @param {*} control
 * @returns
 */
function OnAction(control) {
    var eleId;
    if (typeof control == "object" && arguments.length == 1) { //针对Ribbon的按钮的
        eleId = control.Id;
    } else if (typeof control == "undefined" && arguments.length > 1) { //针对idMso的
        eleId = arguments[1].Id;
        console.log("idMso的ID是：" + eleId)
    } else if (typeof control == "boolean" && arguments.length > 1) { //针对checkbox的
        eleId = arguments[1].Id;
    } else if (typeof control == "number" && arguments.length > 1) { //针对下拉菜单的
        eleId = arguments[2].Id;
    }
    switch (eleId) {
        case "btnShowMsg": {
            const doc = wps.WpsApplication().ActiveDocument
            if (!doc) {
                alert("当前没有打开任何文档")
                return
            }
            alert(doc.Name)
            break;
        }
        case "btnShowDialog":
            //wps.Show*Dialog(GetUrlPath() + "/ui/dialog.html", "这是一个对话框网页", 400 * window.devicePixelRatio, 400 * window.devicePixelRatio, false)
            wps.ShowDialog(GetUrlPath() + "/ui/paper.html", "生成试题", 1200 * window.devicePixelRatio, 700 * window.devicePixelRatio, false)
            break
        case "btnShowTrans":{
            let tsId = wps.PluginStorage.getItem(constom_Enum.TaskPane_ID)
            if (!tsId) {
                let tskpane = wps.CreateTaskPane(GetUrlPath() + "/ui/translate.html")
                let id = tskpane.ID
                wps.PluginStorage.setItem(constom_Enum.TaskPane_ID, id)
                tskpane.Visible = true
            } else {
                let tskpane = wps.GetTaskPane(tsId)
                tskpane.Visible = !tskpane.Visible
            }
            break
        }
        case "btnShowTaskPane": {
            let tsId = wps.PluginStorage.getItem(constom_Enum.TaskPane_ID)
            if (!tsId) {
                let tskpane = wps.CreateTaskPane(GetUrlPath() + "/ui/taskpane.html")
                let id = tskpane.ID
                wps.PluginStorage.setItem(constom_Enum.TaskPane_ID, id)
                tskpane.Visible = true
            } else {
                let tskpane = wps.GetTaskPane(tsId)
                tskpane.Visible = !tskpane.Visible
            }
            break
        }
        case "btnShowLoginDialog": {
            //从WPS加载项的存储空间取值
            let docSelectText = wps.PluginStorage.getItem(constom_Enum.DcoumentSelectionText);
            //获取WPS活动文档已经选中的内容（文字）
            docSelectText = wps.WpsApplication().Selection.Text;
            //将选中的内容存放到WPS加载项的存储空间中，方便各个页面的数据同步
            wps.PluginStorage.setItem(constom_Enum.DcoumentSelectionText, docSelectText);
            //弹出Web对话框，可以是个URL，也可以是WPS加载项的网页
            wps.ShowDialog("www.baidu.com/s?wd=" + docSelectText, "百度搜索", 800 * window.devicePixelRatio, 800 * window.devicePixelRatio, false)
            // // wps.ShowDialog("http://127.0.0.1:8080/browser-integration-wps/", "我是标题，随你心意起名字", 800 * window.devicePixelRatio, 800 * window.devicePixelRatio, true)
            // //增加TabPage
            // wps.TabPages.Add("https://www.baidu.com")
            // //设置任务窗格
            //  let tskpan_login = wps.CreateTaskPane("mail.163.com", "我是个登录框")
            //  let id = tskpan_login.ID
            //  wps.PluginStorage.setItem(constom_Enum.TaskPane_Login_ID, id)
            //  tskpan_login.Visible = true
            //  tskpan_login.Width = 600 * window.devicePixelRatio
            //  tskpan_login.DockPosition = WPS_Enum.msoCTPDockPositionLeft
            //wps.ShowDialog(GetUrlPath() + "/ui/iframe.html", "iframe框", 1000 * window.devicePixelRatio, 1000 * window.devicePixelRatio, false)

            //     var storage = window.localStorage;
            //     //写入a字段
            //     storage["a"] = 1;
            //     //写入b字段
            //     storage.b = 2;
            //     //写入c字段
            //     storage.setItem("c", 3);
            //     console.log(typeof storage["a"]);
            //     console.log(typeof storage["b"]);
            //     console.log(typeof storage["c"]);

            //    alert("我是localStorage的key为a的值：" + localStorage.getItem("a"))



            // wps.ShowDialog("http://127.0.0.1:3888/index2", "我是标题，随你心意起名字", 800 * window.devicePixelRatio, 800 * window.devicePixelRatio, false)
            // wps.ShowDialog(GetUrlPath() + "/index2.html", "我是标题，随你心意起名字", 800 * window.devicePixelRatio, 800 * window.devicePixelRatio, false)
            break
        }
    case "btnShowNewtaskpanel": {
        let tsId = wps.PluginStorage.getItem(constom_Enum.TaskPane_ID)
        if (!tsId) {
            let tskpane = wps.CreateTaskPane(GetUrlPath() + "/ui/iframe2.html")
            // let tskpane = wps.CreateTaskPane("http://127.0.0.1:8080/browser-integration-wps/","我是业务系统的页面")
            let id = tskpane.ID
            wps.PluginStorage.setItem(constom_Enum.TaskPane_ID, id)

            tskpane.Visible = true
        } else {
            let tskpane = wps.GetTaskPane(tsId)
            tskpane.Visible = !tskpane.Visible
        }
        break
    }
    
    case "btnShowMulu": {
        getMulu()
    }
    break
    case "btnOnshowDocumentField": {
        OnshowDocumentField()
    }
    break
    case "btnGetAllReview": {
        OnGetAllReview()
    }
    break
    //内容控件的示例按钮
    case "btnShowControlsTaskPanel": {
        let tsId = wps.PluginStorage.getItem(constom_Enum.TaskPane_ContentControls_ID)
        if (!tsId) {
            let tskpane = wps.CreateTaskPane(GetUrlPath() + "/ui/CreatePaper.html")
            let id = tskpane.ID
            wps.PluginStorage.setItem(constom_Enum.TaskPane_ContentControls_ID, id)
            tskpane.Visible = true
        } else {
            let tskpane = wps.GetTaskPane(tsId)
            tskpane.Visible = !tskpane.Visible
        }
        break
    }
    // case "FileSave": {
    //     alert("FileSave");
    // }
    // case "Copy": {
    //     alert("noCopy");
    // }
    default:
        break
    }
    return true
}

/**
 * dynamicMenu_dmnuShowReview的设置内容的回调函数
 * 给出dynamicMenu最基本的用法：利用文中的修订记录作为menu内容，在文中没有修订记录时显示为空
 * @param {*} control
 * @returns
 */
function dmnuShowReview_getContent(control){
    //  <menu xmlns = "http://schemas.microsoft.com/office/2006/01/customui" >
    //      <button id = "button1" label = "Button 1" / >
    //      <button id = "button2" label = "Button 2" / >
    //      <button id = "button3" label = "Button 3" / >
    //  </menu>
    // 语法可以参考：https://docs.microsoft.com/en-us/openspecs/office_standards/ms-customui/fd0825c7-0f29-4038-8617-54af0dec8c7d
    let sXML = '<menu xmlns = ' + '"http: //schemas.microsoft.com/office/2006/01/customui" >'
    let bXML = ''
    let endXML = '</menu>'
    let l_doc = wps.WpsApplication().ActiveDocument
    for (var index = 1; index <= l_doc.Revisions.Count; index++) {
        bXML += '<button id = ' + '"' + timeConvert(l_doc.Revisions.Item(index).Date).toString()+'_'+index + '"' 
              + 'label = ' + '"' + timeConvert(l_doc.Revisions.Item(index).Date).toLocaleString() +'"' 
              + 'onAction=' + '"dmnBtn_onAction"'
              + '/>'
    }
    if(bXML==''){
        bXML = '<button id = ' + '"dmnu_btn_null"' + 'label = ' + '"空"' + '/>'
    }
    return sXML + bXML + endXML
}
/**
 * dynamicMenu_dmnuShowReview的下拉菜单中的元素点击事件
 * 实现dynamicMenu的内容元素的点击事件
 * @param {*} control
 * @returns
 */
function dmnBtn_onAction(control){
    var eleId;
    if (typeof control == "object" && arguments.length == 1) { //针对Ribbon的按钮的
        eleId = control.Id;
    } else if (typeof control == "undefined" && arguments.length > 1) { //针对idMso的
        eleId = arguments[1].Id;
        // console.log("idMso的ID是：" + eleId)
    } else if (typeof control == "boolean" && arguments.length > 1) { //针对checkbox的
        eleId = arguments[1].Id;
    } else if (typeof control == "number" && arguments.length > 1) { //针对下拉菜单的
        eleId = arguments[2].Id;
    }
    switch (eleId) {
        case "dmnu_btn_null":
            break
        default:
            {
                //弹出对应的修订记录的「作者：修改内容」
                let index = eleId.split("_")[1]
                let l_doc = wps.WpsApplication().ActiveDocument
                alert(l_doc.Revisions.Item(index).Author + "：" + l_doc.Revisions.Item(index).FormatDescription)
            }            
    }
    return true
}

// TODO dropdown，combox复杂控件的支持
/**
 *获取dropdown下拉项的总数
 *
 * @returns
 */
function GetDropDownGetItemCount(control) {
    console.log("GetDropDownGetItemCount" + control)
    let l_doc = wps.WpsApplication().ActiveDocument
    if (l_doc) {
        return l_doc.Revisions.Count
    }
    return 0
}
/**
 * 获取dropdown的每一项的名称
 *
 * @param {*} control
 * @returns
 */
function GetDropDownGetItemLabel(control) {
    console.log("GetDropDownGetItemLabel" + control)
    // console.log("GetDropDownGetItemLabel" + indexx)
    let l_doc = wps.WpsApplication().ActiveDocument
    for (var index = 1; index <= l_doc.Revisions.Count; index++) {
        // console.log(l_doc.Revisions.Item(index))
        return "label" + timeConvert(l_doc.Revisions.Item(index).Date)
    }
}
/**
 * 获取dropdown的每一项的ID
 *
 * @param {*} control
 * @returns
 */
function GetDropDownGetItemID(control) {
    console.log("GetDropDownGetItemID" + control)
    let l_doc = wps.WpsApplication().ActiveDocument
    for (var index = 1; index <= l_doc.Revisions.Count; index++) {
        {
            return l_doc.Revisions.Item(index)
            // return "label" + timeConvert(l_doc.Revisions.Item(index).Date)
        }
    }
}
/**
 * 获取dropdown的每一项的tip
 *
 * @param {*} control
 * @returns
 */
function GetDropDowGetItemScreenTip(control) {
    console.log("GetDropDowGetItemScreenTip" + control)
    let l_doc = wps.WpsApplication().ActiveDocument
    for (var index = 1; index <= l_doc.Revisions.Count; index++) {
        {
            // console.log(l_doc.Revisions.Item(index))
            return "我是" + timeConvert(l_doc.Revisions.Item(index).Date)
        }
    }
}
/**
 * 获取活动文档的所有注释
 *
 */
function OnGetAllReview() {
    let l_doc = wps.WpsApplication().ActiveDocument
    
    wps.ShowDialog(GetUrlPath() + "/test.html", l_doc.Revisions.Count, 600 * window.devicePixelRatio, 600 * window.devicePixelRatio, false)
    for (var index = 1; index <= l_doc.Revisions.Count; index++) {
        console.log(l_doc.Revisions.Item(index))
        console.log(timeConvert(l_doc.Revisions.Item(index).Date))
        
    }
}
/**
 * 通过「智能目录」生成文档的目录结构
 *
 */
function getMulu() {
    var wpsApp = wps.WpsApplication();
    var activeDoc = wpsApp.ActiveDocument;
    var selection = wps.WpsApplication().Selection;
    selection.Range.GetRangeEx().GetHeadings().SmartToRecognize()
    var count = activeDoc.Content.GetRangeEx().GetHeadings().Count
    if (typeof count == "number" && count > 0) {
        var text1 = activeDoc.Content.GetRangeEx().GetHeadings().Item(1).Text
        var text2 = activeDoc.Content.GetRangeEx().GetHeadings().Item(2).Text
        var text3 = activeDoc.Content.GetRangeEx().GetHeadings().Item(4).Text
        alert(count);
        alert(text1);
        alert(text2);
        alert(text3);
    }
}


/**
 * 针对LocalStorage的操作
 *
 */
function alertLocalStorage() {
    // alert("我是localStorage的key为task的值：" + localStorage.getItem("task"));
    alert("我是localStorage的key为a的值：" + localStorage.getItem("a"));
}
/**
 * 控件的图片的获取方法
 *
 * @param {*} control
 * @returns
 */
function GetImage(control) {
    const eleId = control.Id
    switch (eleId) {
        case "btnShowMsg":
            return "images/1.svg"
        case "btnShowDialog":
            return "images/2.svg"
        case "btnShowTaskPane":
            return "images/3.svg"
        case "btnShowLoginDialog":
            return "images/2.svg"
        case "btnShowLoginTaskPane":
            return "images/3.svg"
        default:
            ;
    }
    return "images/newFromTemp.svg"
}


