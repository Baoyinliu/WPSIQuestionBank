/**
 * 统计页面各类书签及内容控件数量
 * ClassNum
 * ChapterNum
 * QuestionTypeNum
 * QuestionNum(来源于内容控件)
 *
 */

 let ClassNum = 0, ChapterNum = 0, QuestionTypeNum = 0, QuestionNum = 0
 let ClassName = "", DocName = ""
 function GetNums(){
     ClassNum = 0, ChapterNum = 0, QuestionTypeNum = 0, QuestionNum = 0
     let l_doc = wps.WpsApplication().ActiveDocument
     let bks = l_doc.Bookmarks
     let l_doc_controls = l_doc.ContentControls
     for(let i=1; i<=bks.Count; i++){
         let BM = bks.Item(i)
         if(BM.Name.indexOf("Class") === 0){
             ClassNum = ClassNum + 1
         }else if(BM.Name.indexOf("CHP") === 0){
             ChapterNum = ChapterNum + 1
         }else if(BM.Name.indexOf("QT") === 0){
             QuestionTypeNum = QuestionTypeNum + 1
         }else if(BM.Name.indexOf("D") === 0){
             QuestionNum = QuestionNum + 1
         }
     }
     // ps.setItem("ClassNum", ClassNum)
     // ps.setItem("ChapterNum", ChapterNum)
     // ps.setItem("QuestionTypeNum", QuestionTypeNum)
     // ps.setItem("QuestionNum", QuestionNum)
 }
 
 
 /**
  * 页面的启动方法
  *
  */
 function onload() {
     wps.ApiEvent.AddApiEventListener("WindowActivate", onDocActiveChange) //当前文档切换后的事件回调通知
     if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
         wps.Enum = WPS_Enum
     }
     getContentControls()
     GetNums()
     GetPaperTitle()
 }

 function GetPaperTitle() {
     
     let l_doc = wps.WpsApplication().ActiveDocument
     if (!l_doc) {
        return
     }
     //获取活动文档的所有内容控件List
     let l_doc_controls = l_doc.ContentControls
     //获取bookMarks
     let bks = l_doc.Bookmarks
     ClassName = bks.Item("Class").Range.Text
     DocName = l_doc.Name
     //插入HTML
     document.getElementById('PaperTitleDiv').innerHTML = ""
     let innerHtml = ""
     innerHtml = "<h2>科目："+ ClassName + "</h2>"
     innerHtml = innerHtml + "<p>试卷来源于文档："+DocName+ "</p>"
     document.getElementById('PaperTitleDiv').innerHTML = innerHtml
 }

 function GetPaper() {
    let QD = document.getElementById("QD").value
    let QT = document.getElementById("QT").value

    wps.PluginStorage.setItem("QD", QD)
    wps.PluginStorage.setItem("QT", QT)

    wps.ShowDialog(GetUrlPath() + "/paper2.html", "试题", 800 * window.devicePixelRatio, 500 * window.devicePixelRatio, false)
    let l_doc = wps.WpsApplication().ActiveDocument
    if (!l_doc) {
       return
    }

    document.getElementById('PaperQuesDiv2').innerHTML = ""
    //获取活动文档的所有内容控件List
    let l_doc_controls = l_doc.ContentControls
    //获取bookMarks
    let bks = l_doc.Bookmarks
    //ChapterIndex
    for(let CHPI = 1; CHPI<=ChapterNum; CHPI++){
        let ChpName = "CHP"+String(CHPI);
        //寻找每章的题型与题目
        let ChpQTypeNum = 0
        var QNum = new Array()
        //遍历当前章节的题型与题目(当前章节的书签数量代表了题型与题目的总数量)
        let ChpBks = bks.Item(ChpName).Range.Bookmarks
        //第一遍遍历得到题型数目
        for(let i = 1; i<=ChpBks.Count; i++){
            if(ChpBks.Item(i).Name.indexOf("QT") == 0){
                ChpQTypeNum = ChpQTypeNum + 1
                QNum[ChpQTypeNum] = 0
            }
        }
        //第二遍遍历得到对应题型下的题目数量
        for(let i = 1; i<=ChpBks.Count; i++){
            if(ChpBks.Item(i).Name.indexOf("D") == 0){//是题目准备计数
                for(let j=1; j<=ChpQTypeNum; j++){//遍历题型
                    if(ChpBks.Item(i).Name.indexOf("QT"+String(j)) != -1){//题目属于对应题型
                        QNum[j] =  QNum[j] + 1
                    }
                }
            }
        }
        //console.log(QNum)
        // console.log(bks.Item(ChpName).Range.Paragraphs.Item(1).Range.Text)
        //插入HTML
        //获得约束条件
        
        InsertTest(CHPI, ChpQTypeNum, QNum, QD, QT)
    }
 }

 //生成试题
 function InsertTest(CHPI, ChpQTypeNum, QNum, QD, QT){
    let l_doc = wps.WpsApplication().ActiveDocument
    if (!l_doc) {
       return
    }
    //获取活动文档的所有内容控件List
    let l_doc_controls = l_doc.ContentControls
    //获取对应章节bookMarks
    let ChpBookName = "CHP"+String(CHPI)//当前章节的bookmarkName
    let ChpBks = l_doc.Bookmarks.Item(ChpBookName).Range.Bookmarks//章节范围bookmarks
    let ChpName = ChpBks.Item(ChpBookName).Range.Paragraphs.Item(1).Range.Text//章节名字


     //插入HTML
     let PreHTML = document.getElementById('PaperQuesDiv2').innerHTML//保存前面的html语句
     let innerHtml = ""
     innerHtml = PreHTML + "<h3>章节："+ ChpName + "</h3>"
     document.getElementById('PaperQuesDiv2').innerHTML = innerHtml
     for(let i=1; i<=ChpQTypeNum; i++){//遍历题型
        let TypeBookName = "QT"+String(i)+"In"+ ChpBookName
        let TypeName = ChpBks.Item(TypeBookName).Range.Paragraphs.Item(1).Range.Text//题型名字
        //插入HTML(题型)
        PreHTML = document.getElementById('PaperQuesDiv2').innerHTML//保存前面的html语句
        let innerHtml = ""
        innerHtml = PreHTML + "<h4>"+ TypeName + "</h4>"
        document.getElementById('PaperQuesDiv2').innerHTML = innerHtml

        for(let j=1; j<=QNum[i]; j++){//遍历i题型下的所有题目
            let QuestionBookName = "D"+String(j)+"In"+ TypeBookName
            console.log(QuestionBookName)

            let AllContent = ChpBks.Item(QuestionBookName).Range.ContentControls.Item(1)
            let QRange = AllContent.Range
            QRange.End = AllContent.Range.ContentControls.Item(1).Range.Start

            let QuestionText = QRange.Text
            let AText = AllContent.Range.ContentControls.Item(1).Range.Text
            let DText = AllContent.Range.ContentControls.Item(2).Range.Text
            let TText = AllContent.Range.ContentControls.Item(3).Range.Text
            // QuestionText = QuestionText.slice(0, QuestionText.length-1)
            // QuestionText = QuestionText+"\n"+ChpBks.Item(QuestionBookName).Range.ContentControls.Item(2).Range.Text
            
            if(QD == '0' || DText.indexOf(QD) != -1)
            {
                if(QT == '0' || TText.indexOf(QT) != -1)
                {
                    //插入HTML(题目)
                    console.log(QuestionText)
                    PreHTML = document.getElementById('PaperQuesDiv2').innerHTML//保存前面的html语句
                    innerHtml = ""
                    innerHtml = PreHTML +"<div class= \"PaperQues\" > <textarea class='TA' cols=50 rows=5>"+ QuestionText +"</textarea>" 
                    +"<textarea class='TA' cols=50 rows=5>"+ "填写答案：" +"</textarea>" 
                    +
                    "</div>"
                    document.getElementById('PaperQuesDiv2').innerHTML = innerHtml
                }
            }
        }
        //console.log(ChpBks.Item(TypeBookName).Range)
        console.log("dsadasdasd")
    }

 }


 /**
 * 当前文档切换到其它文档时， 重新获取新的文档的内容控件
 *
 * @param {*} doc
 * @param {*} docwin
 */
function onDocActiveChange(doc, docwin) {
    clearContentControls()
    getContentControls()
    GetNums()
}

/**
 * 清空内容控件的列表
 *
 */
function clearContentControls() {
    document.getElementById('contentControlsDiv').innerHTML = ''
}
/**
 * 获取内容控件的列表
 *
 */
function getContentControls() {
    //设置活动文档对象
    let l_doc = wps.WpsApplication().ActiveDocument
    if (!l_doc) {
        return
    }
    GetNums()

    document.getElementById('contentControlsDiv').innerHTML = ""
    let innerHtml = ""
    innerHtml = "<table border='1' class='altrowstable' id='alternatecolor' ><tr><th>章节</th><th>题号</th><th>题目内容</th><th>答案</th><th>难度</th><th>熟练度</th><th>操作</th></tr>"
    let bks = l_doc.Bookmarks//获取书签
    for(let BI=1; BI<=ChapterNum ; BI++){
        let BM = bks.Item(BI)
        //获取活动文档的所有内容控件List
        let l_doc_controls = BM.Range.ContentControls
        let l_doc_controls_count = l_doc_controls.Count//数量内容控件

        for (let i = 1; i <= l_doc_controls_count; ++i) {
            let l_doc_control = l_doc_controls.Item(i)
            let l_doc_control_title = l_doc_control.Title
            if(l_doc_control_title.indexOf("第") == 0)
            {
                let CHText = "第"+ String(BI) +"章"
                let tempRange = l_doc_control.Range
                tempRange.End = l_doc_control.Range.ContentControls.Item(1).Range.Start
                let l_doc_control_text = tempRange.Text
                let AnswerText = l_doc_control.Range.ContentControls.Item(1).Range.Text
                let DText = l_doc_control.Range.ContentControls.Item(2).Range.Text
                let TText = l_doc_control.Range.ContentControls.Item(3).Range.Text
                l_doc_control_text 
                if(l_doc_control.Type == 2) {
                    innerHtml = innerHtml + "<tr>" +
                    "<td width='15%'><li onclick=\"contentControlClick(" + i + ")\")>" + l_doc_control_title + "</li></td>" +
                    "<td><textarea class='TA' id='cc_" + i + "' cols=\"40\" rows=\"5\"  >" +l_doc_control_text+ "</textarea></td>" +
                    "<td width='8%'><button class='btn' onclick=\"ccInsertImg(" + i + ")\")>" + "插入图片" + "</button></td>" +
                    "</tr>"
                }else{
                    innerHtml = innerHtml + "<tr>" +
                    "<td width='15%'>" + CHText + "</td>"+
                    "<td width='15%'><li onclick=\"contentControlClick(" + i + ")\")>" + l_doc_control_title + "</li></td>" +
                    "<td width='30%'><textarea class='TA' id='cc_" + i +"_1" + "' cols=\"35\" rows=\"3\"  >" +l_doc_control_text+ "</textarea></td>" +
                    "<td width='30%'><textarea class='TA' id='cc_" + i +"_2"+ "' cols=\"30\" rows=\"3\"  >" +AnswerText+ "</textarea></td>" +
                    "<td width='30%'><textarea class='TA' id='cc_" + i +"_3"+ "' cols=\"10\" rows=\"3\"  >" +DText+ "</textarea></td>" +
                    "<td width='30%'><textarea class='TA' id='cc_" + i +"_4"+ "' cols=\"10\" rows=\"3\"  >" +TText+ "</textarea></td>" +
                    "<td><button class='btn' onclick=\"ccSetValue(" + i + ")\")>" + "设置" + "</button> " +
                    "<button class='btn' onclick=\"ccGetTxt(" + i + ")\")>" + "提取文本" + "</button> " +
                    "<button class='btn' onclick=\"ccSetFont(" + i + ")\")>" + "设置格式" + "</button> </td>" +
                    "</tr>"
                }
                // innerHtml = innerHtml + "<li class='li' onclick=\"bookMarkClick('" + l_doc_control_title + "')\">" + l_doc_control_title +
                //     "</li>"
            }
        }
    }
    innerHtml = innerHtml + "</table>"
    document.getElementById('contentControlsDiv').innerHTML = innerHtml
    altRows("alternatecolor")
    
}

function altRows(id){
    if(document.getElementsByTagName){
        var table = document.getElementById(id);

        var rows = table.getElementsByTagName("tr");

        for(i = 0; i < rows.length; i++){
            if(i % 2 == 0){
                rows[i].className = "evenrowcolor";
            }else{
                rows[i].className = "oddrowcolor";
            }
        }
    }
}
/**
 * 点击内容控件标题做定位
 *
 * @param {*} bookMarkName
 */
function contentControlClick(ccIndex) {
    let l_doc = wps.WpsApplication().ActiveDocument
    if (!l_doc)
        return
    let l_doc_control = l_doc.ContentControls.Item(ccIndex)
    if (l_doc_control)
        l_doc_control.Range.Select()
}
/**
 * 点击内容控件设置做内容替换
 *
 * @param {*} ccIndex
 */
function ccSetValue(ccIndex) {
    let l_doc = wps.WpsApplication().ActiveDocument

    let QT = document.getElementById("cc_" + ccIndex +"_1").value
    let AT = document.getElementById("cc_" + ccIndex +"_2").value
    let DT = document.getElementById("cc_" + ccIndex +"_3").value
    let TT = document.getElementById("cc_" + ccIndex +"_4").value

    if (!l_doc)
        return
    let l_doc_control = l_doc.ContentControls.Item(ccIndex)

    if (l_doc_control){
        let QRange = l_doc_control.Range
        QRange.End = l_doc_control.Range.ContentControls.Item(1).Range.Start
        QRange.Text = QT

        l_doc_control.Range.ContentControls.Item(1).Range.Text = AT   
        l_doc_control.Range.ContentControls.Item(2).Range.Text = DT   
        l_doc_control.Range.ContentControls.Item(3).Range.Text = TT        
    }

    console.log(l_doc_control.Range)
}