//在后续的wps版本中，wps的所有枚举值都会通过wps.Enum对象来自动支持，现阶段先人工定义
var WPS_Enum = {
    msoCTPDockPositionLeft:0,
    msoCTPDockPositionRight:2
}

function GetUrlPath() {
    let e = document.location.toString()
    return -1!=(e=decodeURI(e)).indexOf("/")&&(e=e.substring(0,e.lastIndexOf("/"))),e
}

var constom_Enum = {
    /**
     * 当前活动文档选中的文字
     */
    DcoumentSelectionText: "DcoumentSelectionText",
    /**
     * 按钮的控制的常量
     */
    EnableFlag: "EnableFlag",
    TaskPane_ID: "TaskPane_ID",
    TaskPane_Login_ID: "TaskPane_Login_ID",
    TaskPane_ContentControls_ID: "TaskPane_ContentControls_ID"
}

/**
 * 作用：创建一个任务窗格，加载指定URL
 * @param {*} taskPanUrl ：任务窗格要加载的URL
 * @param {*} taskPanName ：任务窗格的名称
 * @param {*} taskPanPosition ：任务窗格的位置：right:居右，left:居左
 * @param {*} taskPanWidth ：任务窗格的宽度，px
 */
function creatTaskPanforTranslation(ENum_taskpaneid, taskPanUrl, taskPanName, taskPanPosition, taskPanWidth) {
    let id = wps.PluginStorage.getItem(ENum_taskpaneid);
    let tp;
    if (id) {
        tp = wps.GetTaskPane(id);
        //每次都初始化
        // tp.Delete();
    } else {
        tp = wps.CreateTaskPane(taskPanUrl, taskPanName);
        if (tp) {
            wps.PluginStorage.setItem(ENum_taskpaneid, tp.ID);
        }
    }
    switch (taskPanPosition) {
        case "right":
            tp.DockPosition = WPS_Enum.msoCTPDockPositionRight; //这里可以设置taskapne是在左边还是右边
            break;
        case "left":
            tp.DockPosition = WPS_Enum.msoCTPDockPositionLeft;
            break;
        default:
            tp.DockPosition = WPS_Enum.msoCTPDockPositionRight;
            break;
            // case "bottom":
            //     tp.DockPosition = WPS_Enum.msoCTPDockPosition;
    }
    tp.Width = taskPanWidth;
    tp.Visible = true;
}
