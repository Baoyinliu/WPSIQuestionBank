/**
 * 统计页面各类书签及内容控件数量
 * ClassNum
 * ChapterNum
 * QuestionTypeNum
 * QuestionNum(来源于内容控件)
 *
 */

let ClassNum = 0, ChapterNum = 0, QuestionTypeNum = 0, QuestionNum = 0
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
    //获取活动文档的所有内容控件List
    let l_doc_controls = l_doc.ContentControls

    document.getElementById('contentControlsDiv').innerHTML = ""
    let innerHtml = ""
    innerHtml = "<table border='1'><tr><th>控件名称</th><th>控件内容</th><th>操作</th></tr>"

    let l_doc_controls_count = l_doc_controls.Count//数量内容控件

    for (let i = 1; i <= l_doc_controls_count; ++i) {
        let l_doc_control = l_doc_controls.Item(i)
        let l_doc_control_title = l_doc_control.Title
        let l_doc_control_text = l_doc_control.Range.Text
        if(l_doc_control.Type == 2) {
            innerHtml = innerHtml + "<tr>" +
            "<td width='15%'><li onclick=\"contentControlClick(" + i + ")\")>" + l_doc_control_title + "</li></td>" +
            "<td><input  id='cc_" + i + "' value=" + l_doc_control_text + " style='width:100%;height:100%;'></input></td>" +
            "<td width='8%'><button onclick=\"ccInsertImg(" + i + ")\")>" + "插入图片" + "</button></td>" +
            "</tr>"
        }else{
            innerHtml = innerHtml + "<tr>" +
            "<td width='15%'><li onclick=\"contentControlClick(" + i + ")\")>" + l_doc_control_title + "</li></td>" +
            "<td><input  id='cc_" + i + "' value=" + l_doc_control_text + " style='width:100%;height:100%;'></input></td>" +
            "<td width='8%'><button onclick=\"ccSetValue(" + i + ")\")>" + "设置" + "</button></td>" +
            "<td width='8%'><button onclick=\"ccGetTxt(" + i + ")\")>" + "提取文本" + "</button></td>" +
            "<td width='8%'><button onclick=\"ccSetFont(" + i + ")\")>" + "设置格式" + "</button></td>" +
            "</tr>"
        }
        // innerHtml = innerHtml + "<li class='li' onclick=\"bookMarkClick('" + l_doc_control_title + "')\">" + l_doc_control_title +
        //     "</li>"
    }
    innerHtml = innerHtml + "</table>"

    document.getElementById('contentControlsDiv').innerHTML = innerHtml
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

    let l_doc_cc_value = document.getElementById("cc_" + ccIndex).value

    if (!l_doc)
        return
    let l_doc_control = l_doc.ContentControls.Item(ccIndex)

    if (l_doc_control)
        l_doc_control.Range.Text = l_doc_cc_value

    console.log(l_doc_control.Range)
}


/* 插入文本内容控件并设置标题 */
function ccInsertTxtControl() {
    let l_doc = wps.WpsApplication().ActiveDocument
	let txt = l_doc.ContentControls.Add(wps.wdContentControlText);
	txt.Title = "姓名2号"

    
}

function ccInsertImgControl() {
    let l_doc = wps.WpsApplication().ActiveDocument

	let img = l_doc.ContentControls.Add(2);
	img.Title = "证件照2号"

}

function ccInsertImg(ccIndex) {
    let l_doc = wps.WpsApplication().ActiveDocument
    let ImgFile = document.getElementById("cc_" + ccIndex).value
    if(!l_doc)
        return
    let l_doc_control = l_doc.ContentControls.Item(ccIndex)
    if(l_doc_control){
        console.log(l_doc_control.Range.InlineShapes)
        l_doc_control.Range.InlineShapes.AddPicture(ImgFile)
    }
}


function ccGetTxt(ccIndex) {
    let l_doc = wps.WpsApplication().ActiveDocument
    let MyTxt = ""
    if(!l_doc)
        return
    let l_doc_control = l_doc.ContentControls.Item(ccIndex)
    if(l_doc_control){
        MyTxt = l_doc_control.Range.Text
    }
    let MyRange = l_doc.Range(l_doc.Paragraphs.Item(1).Start, l_doc.Paragraphs.Item(1).End - 1)
    MyRange.InsertAfter(MyTxt)
}


function ccGetAll(){
    let MyTxt="Sum: "
    let l_doc = wps.WpsApplication().ActiveDocument
    if (!l_doc) {
        return
    }
    //获取活动文档的所有内容控件List
    let l_doc_controls = l_doc.ContentControls
    let ContenControlNum = l_doc_controls.Count
    for (let i = 1; i <= ContenControlNum; ++i) {
        let l_doc_control = l_doc_controls.Item(i)
        let l_doc_control_title = l_doc_control.Title
        let l_doc_control_text = l_doc_control.Range.Text
        if(l_doc_control.Type != 2)
            MyTxt = MyTxt + l_doc_control_title + ":" + l_doc_control_text
    }
    let MyRange = l_doc.Range(l_doc.Paragraphs.Item(1).Start, l_doc.Paragraphs.Item(1).End - 1)
    MyRange.InsertAfter(MyTxt)
}


function ccSetFont(ccIndex) {
    let l_doc = wps.WpsApplication().ActiveDocument
    if(!l_doc)
        return
    let l_doc_control = l_doc.ContentControls.Item(ccIndex)
    if(l_doc_control){
        let MyRange = l_doc_control.Range
        MyRange.Font.Name = "Arial"
        MyRange.Font.Bold = true
        MyRange.Font.ColorIndex = 	128
    }
    
}


function ccMerge(){
    let MyTxt="Merge: "
    let docs = wps.WpsApplication().Documents
    if(!docs)
        return
    for (let i = 1; i <= docs.Count; ++i) {
        docs.Item(i).Activate()
        let l_doc = wps.WpsApplication().ActiveDocument
        if(!l_doc)
            continue
        //获取活动文档的所有内容控件List
        let l_doc_controls = l_doc.ContentControls
        let ContenControlNum = l_doc_controls.Count
        for (let i = 1; i <= ContenControlNum; ++i) {
            let l_doc_control = l_doc_controls.Item(i)
            let l_doc_control_title = l_doc_control.Title
            let l_doc_control_text = l_doc_control.Range.Text
            if(l_doc_control.Type != 2)
                MyTxt = MyTxt + l_doc_control_title + ":" + l_doc_control_text
        } 
    }

    wps.WpsApplication().Documents.Add()
    docs.Item(docs.Count+1).Activate()
    
    let new_doc = wps.WpsApplication().ActiveDocument
    let newRange = new_doc.Range(new_doc.Paragraphs.Item(1).Start, new_doc.Paragraphs.Item(1).End - 1)
    newRange.InsertAfter(MyTxt)
}


function InsertClass() {

    let l_doc = wps.WpsApplication().ActiveDocument
    //插入新的段落作为科目
    l_doc.Paragraphs.Add(l_doc.Paragraphs.Item(1).Range)
    l_doc.Paragraphs.Item(1).Range.InsertBefore("语文")
    //设置为一级标题
    l_doc.Paragraphs.Item(1).Style = 3
    //科目采用Class作为标记
    l_doc.Bookmarks.Add("Class", l_doc.Paragraphs.Item(1).Range)
    ClassNum = ClassNum + 1    
    
}

/**
 * 插入章节
 *
 * 
 */
function InsertChapter() {

    let l_doc = wps.WpsApplication().ActiveDocument
    let bks = l_doc.Bookmarks//获取书签
    //插入题型如果未插入科目报错
    if(ClassNum === 0) {
        alert("未设置科目！")
        return
    }
    let CurRange = l_doc.Range()
    //如果是第一次插入章节
    if(ChapterNum == 0) {
         //选中
        bks.Item("Class").Select();
        //范围来到科目书签的下面一行
        CurRange = l_doc.Range().GoTo(-1, 1, 1)
        CurRange = CurRange.GoTo(3, 2, 1)
    }else if(ChapterNum != 0) {
        //范围来到最后一个章节的最后一个题型的最后一道题下面一行
        bks.Item("CHP"+String(ChapterNum)).Select();
        CurRange = bks.Item("CHP"+String(ChapterNum)).Range
        CurRange = CurRange.GoTo(3, 2, 1)
    }
    //插入章节内容
    ChapterNum = ChapterNum + 1
    alert(ChapterNum)
    CurRange.InsertParagraph()
    CurRange.InsertBefore("章节"+String(ChapterNum))
    //插入bookmark
    l_doc.Bookmarks.Add("CHP"+String(ChapterNum), CurRange)
    //设置二级标题
    CurRange.Paragraphs.Item(1).Style = 4 
}
/**
 * 插入题型
 *
 * 
 */
function InsertQuestionType() {
    let At = 1
    let CurQuestionType = 0
    Chaptername = "CHP"+String(At)
    let l_doc = wps.WpsApplication().ActiveDocument
    let bks = l_doc.Bookmarks//获取书签
    let ChapterBk = bks.Item("CHP"+String(At))
    let ChapterRange = ChapterBk.Range//定位章节范围
    //插入题型如果未插入科目报错
    if(ChapterNum == 0) {
        alert("未设置章节！")
        return
    }
    //定义变量
    let CurRange = l_doc.Range()
    let CurBook = ChapterRange.Bookmarks.Item(1)
    //如果是第一次插入题型
    if(ChapterRange.Bookmarks.Count == 1) {
        //选中章节
        CurBook = ChapterRange.Bookmarks.Item(Chaptername)
        CurBook.Range.Select();
        //范围来到章节下面一行
        CurRange = CurBook.Range
        CurRange = CurRange.GoTo(3, 2, 1)
    }else{
        //获得一下题型数量
        for(let i=1; i<=ChapterRange.Bookmarks.Count; i++){
            let BM = ChapterRange.Bookmarks.Item(i)
            if(BM.Name.indexOf("QT") === 0){
                console.log(BM.Name)
                CurQuestionType = CurQuestionType + 1
            }
        }
        //来到章节中的最后一个书签
        CurBook = ChapterRange.Bookmarks.Item("QT"+String(CurQuestionType)+"In"+Chaptername)
        console.log("QT"+String(CurQuestionType)+"In"+Chaptername)
        CurBook.Range.Select();
        //范围来到最后一个题型下面一行
        CurRange = CurBook.Range
        CurRange = CurRange.GoTo(3, 2, 1)
    }
    //插入题型内容
    QuestionTypeNum = QuestionTypeNum + 1
    CurRange.InsertParagraph()
    CurRange.InsertBefore("题型"+String(CurQuestionType+1))
    //插入bookmark
    l_doc.Bookmarks.Add("QT"+String(CurQuestionType+1)+"In"+Chaptername, CurRange)
    //设置三级标题
    CurRange.Paragraphs.Item(1).Style = 5

    //更新章节书签***********
    let ChNewRange = ChapterRange
    ChNewRange.End = CurRange.End
    l_doc.Bookmarks.Add(Chaptername, ChNewRange)
    //还需要更新下一个章节书签
    if(ChapterNum >=  At+1){
        ChNewRange = bks.Item("CHP"+String(At+1)).Range
        ChNewRange.Start = CurRange.End
        l_doc.Bookmarks.Add("CHP"+String(At+1), ChNewRange)
    }
}

/**
 *  获得书签索引
 *
 */

function GetBookMarkIndex(BName) {
    let l_doc = wps.WpsApplication().ActiveDocument
    let bks = l_doc.Bookmarks//获取书签
    if(bks.Exists(BName) === false) {
        alert("书签不存在")
        return -1
    }
    for(let i = 1;i <= bks.Count; i++){
        if(bks.Item(i).Name === BName)
            return i
    }
}

/**
 * 插入题目
 *
 * 
 */
function InsertQuestion() {
    let l_doc = wps.WpsApplication().ActiveDocument
    let bks = l_doc.Bookmarks//获取书签
    //插入bookMark Range内最后一行新段落，并更新bookMark范围，副作用是下一个bookMark Range会改变
    let At = 1//章节序号

    //未设置题型
    if(bks.Item("CHP"+String(At)).Count == 1) {
        alert("未设置题型！")
        return
    }

    let Type = 1//题型序号
    let BNameChapter = "CHP"+String(At)
    let BNameType = "QT" + String(Type) + "In" + BNameChapter
    let CurBook = bks.Item(BNameType)
    let CurRange = CurBook.Range

    //插入题目内容
    //得到题号
    let Qnum = 1
    for(let i=1; i<= CurRange.Bookmarks.Count; i++){
        if(CurRange.Bookmarks.Item(i).Name.indexOf("D") != -1 && CurRange.Bookmarks.Item(i).Name.indexOf(BNameType) != -1){
            Qnum = Qnum+1
        }
    }
    let Q = "第"+ String(Qnum) +"题\t\n"//题目标号玄学占位符
    let A = "第"+ String(Qnum) +"题答案\n"//答案标号
    CurRange.Paragraphs.Add()
    CurRange.InsertAfter(Q)
    //设置题目正文
    CurRange.Paragraphs.Item(CurRange.Paragraphs.Count-1).Style = 1
    CurRange.Paragraphs.Item(CurRange.Paragraphs.Count).Style = 1
    //更新题型bookMark
    l_doc.Bookmarks.Add(BNameType, CurRange)  
    //设置题目bookMark*****************
    let QRange = CurBook.Range
    //位置需要跳过第*题
    QRange.Start = CurBook.Range.End+Q.length
    QRange.End = QRange.Start
    //选择好区域以插入内容控件
    QRange.Select()
    //Question and Answer 两个控件
    let MyQuestion = "到底tm的能不能换行，到底能不能插入问题？？，啊啊啊啊啊啊啊啊啊啊啊啊啊啊tell me ok()"
    let QControl = InsertTest(Q, MyQuestion)//Question内容控件Range
    //QControl.Range.InsertParagraph()//插入段落相当于换行
    QControl.Range.Paragraphs.Add()
    //Answer 控件的位置要向下移动
    QRange.Start = QRange.End+MyQuestion.length+3//玄学占位符
    QRange.End = QRange.Start
    QRange.Select()
    let AControl = InsertTest(A, "A、行   B、不行  C、管你行不行  D、滚蛋")
    let MyRange = QControl.Range
    MyRange.End = AControl.Range.End
    l_doc.Bookmarks.Add("D"+String(Qnum) + "In"+ BNameType, MyRange)
    //修正Question内容控件Range
    //QControl.Range.Text = QControl.Range.Text.slice(0, QControl.Range.Text.length-1)
    //******************************************************************************* */
    

    //修正下一bookMarkRange
    let NextBName = "QT" + String(Type+1) + "In" + BNameChapter
    let NextBookIndex = GetBookMarkIndex(NextBName)
    if(NextBookIndex === -1){
        //当前就是最后一个bookmark 无需修改
        //更新章节书签***********
        let ChapterBk = bks.Item("CHP"+String(At))
        let ChapterRange = ChapterBk.Range//定位章节范围
        let ChNewRange = ChapterRange
        ChNewRange.End = CurRange.End
        ChNewRange.Select()
        l_doc.Bookmarks.Add(Chaptername, ChNewRange)
        if(ChapterNum >=  At+1){
            ChNewRange = bks.Item("CHP"+String(At+1)).Range
            ChNewRange.Start = CurRange.End
            l_doc.Bookmarks.Add("CHP"+String(At+1), ChNewRange)
        }
    }

    //正确的bookmark在上个cur最后一行
    let NextBook = bks.Item(NextBookIndex)
    let NextRange = NextBook.Range
    NextRange.Start = CurRange.End
    //NextRange.Select()
    l_doc.Bookmarks.Add(NextBName, NextRange)

}

////设置标题等级
// function SetHeading(){
//     let l_doc = wps.WpsApplication().ActiveDocument
//     //一级标题：3 二级标题：4 三级标题：5 四级标题：6
//     wps.WpsApplication().Selection.Paragraphs.Item(1).Style = 6
// }

/* 插入文本内容控件并设置标题 */
function InsertTest(T, Text) {
    let l_doc = wps.WpsApplication().ActiveDocument
	let Qtxt = l_doc.ContentControls.Add(1);
	Qtxt.Title = T
    Qtxt.MultiLine = true
    Qtxt.Range.Text = Text
    return Qtxt
}
// function SetTest(Q) {
//     let l_doc = wps.WpsApplication().ActiveDocument
// 	let Qtxt = l_doc.ContentControls.Item(1);
//     console.log(Qtxt)
//     Qtxt.Range.Text = "第一题这到底能不能换行换行换行换行换行换行换行换行啊啊啊啊啊啊能不能到底？？？？？"

// }

function test() {
    let l_doc = wps.WpsApplication().ActiveDocument
    l_doc.Bookmarks.Item("QT3InCHP1").Range.Select()
    let se = wps.WpsApplication().Selection
    console.log(se.Range.Bookmarks.Count)
    for(let i =1;i<=se.Range.Bookmarks.Count;i++){
        let a = se.Range.Bookmarks.Item(i)
        console.log(a.Name)
    }
    //console.log(GetBookMarkIndex("QT3InCHP1"))
}