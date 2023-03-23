/**
 * 统计页面各类书签及内容控件数量
 * ClassNum
 * ChapterNum
 * QuestionTypeNum
 * QuestionNum(来源于内容控件)
 *
 */
 let ClassName = "", DocName = ""
 let ClassNum = 0, ChapterNum = 0, QuestionTypeNum = 0, QuestionNum = 0
 let Chapters = new Array()
 let QTypes = new Array()
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
             Chapters[ChapterNum] = new Array()
             Chapters[ChapterNum][0] = BM.Name//书签名
             Chapters[ChapterNum][1] = bks.Item(i).Range.Paragraphs.Item(1).Range.Text//真名
             Chapters[BM.Name] = ChapterNum
         }else if(BM.Name.indexOf("QT") === 0){
             QuestionTypeNum = QuestionTypeNum + 1
             QTypes[BM.Name] = bks.Item(BM.Name).Range.Paragraphs.Item(1).Range.Text//题型名字
         }else if(BM.Name.indexOf("D") === 0){
             QuestionNum = QuestionNum + 1
         }
     }
     if(ClassNum != 0)
        ClassName = bks.Item("Class").Range.Text
     DocName = l_doc.Name

     if(ClassNum === 0){//科目为空
        document.getElementById('ClassName').value = "未设置科目"
     }else{
        document.getElementById('ClassName').value = ClassName
     }

     //更新页面信息
 }

 function UpdatePage(){
     //插入题型部分*****************************************************************************
     document.getElementById('InsertChp').innerHTML = ""
     let innerHTML = ""
     let tempHTML ="<select id=\"ChapterSelect\">"
     for(let i=1; i<=ChapterNum; i++){
         //console.log(Chapters[i][1])
         //实际value是章节的bookName
         tempHTML = tempHTML + "<option value = "+ Chapters[i][0] + ">"+Chapters[i][1]+"</option>"
     }
     tempHTML = tempHTML + "</select>"
     document.getElementById('InsertChp').innerHTML = innerHTML + tempHTML
     document.getElementById('InsertQType').innerHTML = " <button class='btn' onclick = \"InsertQuestionType()\">插入</button class='btn'> "
     
     //插入题目部分*******************************************************************************
     //选择章节--------------------
     document.getElementById('InsertQASelectChp').innerHTML = ""
     tempHTML ="<select id=\"InsertQAChapterSelect\">"
     for(let i=1; i<=ChapterNum; i++){
         //console.log(Chapters[i][1])
         //实际value是章节的bookName
         tempHTML = tempHTML + "<option value = "+ Chapters[i][0] + ">"+Chapters[i][1]+"</option>"
     }
     tempHTML = tempHTML + "</select>"
     document.getElementById('InsertQASelectChp').innerHTML = tempHTML
     document.getElementById('InsertQASelectChp').innerHTML += " <button  class='btn' onclick = \"InsertQASelectChp()\">确定</button class='btn'> "




     //插入题干-------------------
     document.getElementById('InsertQuestion').innerHTML = ""
     let QHTML =  "<select id=\"QInputType\">" +"<option value = \"text\">文本</option> "
     +"<option value = \"img\">图片</option>" 
     +"</select>"
     //QHTML += "<input type=\"text\" id=\"QInput\">"
     QHTML += " <button class='btn' onclick = \"QInputType()\">确定</button class='btn'> "
     document.getElementById('InsertQuestion').innerHTML = QHTML 

     //插入答案-------------------
     document.getElementById('InsertQAnswer').innerHTML = ""
     let QAHTML =  "<select id=\"QAInputType\">" +"<option value = \"text\">文本</option> "
     +"<option value = \"img\">图片</option>" 
     +"</select>"
     //QHTML += "<input type=\"text\" id=\"QInput\">"
     QAHTML += " <button class='btn' onclick = \"QAInputType()\">确定</button class='btn'> "
     document.getElementById('InsertQAnswer').innerHTML = QAHTML  
 }
 
 
//  setInterval(function(){   
//      document.getElementById('InsertQuestion').innerHTML = ""
//      let QHTML =  "<select id=\"QInputType\">" +"<option value = \"text\">文本</option> "
//      +"<option value = \"img\">图片</option>" 
//      +"</select>"
//      if(document.getElementById('QInputType').value == "text")
//         QHTML += "<input type=\"text\" id=\"QInput\">"
//      document.getElementById('InsertQuestion').innerHTML = QHTML 
//   }, 30);
 
 /**
  * 页面的启动方法
  *
  */
 function onload() {
     wps.ApiEvent.AddApiEventListener("WindowActivate", onDocActiveChange) //当前文档切换后的事件回调通知
     if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
         wps.Enum = WPS_Enum
     }
     GetNums()
     UpdatePage()
 }
 
 /**
  * 当前文档切换到其它文档时， 重新获取新的文档的内容控件
  *
  * @param {*} doc
  * @param {*} docwin
  */
 function onDocActiveChange(doc, docwin) {
     GetNums()
 }
 
 
  /**
  * 设置科目
  *
  * 
  */
 function SetClass() {
 
     let l_doc = wps.WpsApplication().ActiveDocument
     if(ClassNum == 0){//第一次设置科目
        //插入新的段落作为科目
        l_doc.Paragraphs.Add(l_doc.Paragraphs.Item(1).Range)
        l_doc.Paragraphs.Item(1).Range.InsertBefore(document.getElementById('ClassName').value)
        //设置为一级标题
        l_doc.Paragraphs.Item(1).Style = 3
        //科目采用Class作为标记
        l_doc.Bookmarks.Add("Class", l_doc.Paragraphs.Item(1).Range)
        ClassNum = ClassNum + 1    
     }
     else{
        //无语非得这样才能保持格式不变更改
        l_doc.Bookmarks.Item("Class").Range.InsertParagraphBefore()
        l_doc.Bookmarks.Item("Class").Range.InsertBefore(document.getElementById('ClassName').value)
        l_doc.Bookmarks.Item("Class").Range.Paragraphs.Item(2).Range.Delete()       
     }
     
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
         //CurRange.InsertParagraphAfter()
         CurRange.InsertAfter("here")
         CurRange = CurRange.Paragraphs.Item(CurRange.Paragraphs.Count).Range
         CurRange.Select()
     }
     //插入章节内容
     ChapterNum = ChapterNum + 1
     alert(ChapterNum)
     CurRange.InsertParagraph()

     //获取自定义章节名称
     CurRange.InsertBefore("章节"+String(ChapterNum)+"--"+document.getElementById('ChapterName').value)
     //插入bookmark
     l_doc.Bookmarks.Add("CHP"+String(ChapterNum), CurRange)
     //设置二级标题
     CurRange.Paragraphs.Item(1).Style = 4 

     GetNums()//刷新用
 }
 /**
  * 插入题型
  *
  * 
  */
 function InsertQuestionType() {
     
     let Chaptername =  document.getElementById('ChapterSelect').value
     let At = Chapters[Chaptername]//章节标号
     let CurQTypeNum = 0//题目类型数量
     let l_doc = wps.WpsApplication().ActiveDocument
     let bks = l_doc.Bookmarks//获取书签
     let ChapterBk = bks.Item(Chaptername)
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
                 //console.log(BM.Name)
                 CurQTypeNum = CurQTypeNum + 1
             }
         }
         //来到章节中的最后一个书签
         CurBook = ChapterRange.Bookmarks.Item("QT"+String(CurQTypeNum)+"In"+Chaptername)
         console.log("QT"+String(CurQTypeNum)+"In"+Chaptername)
         CurBook.Range.Select();
         //范围来到最后一个题型下面一行
         CurRange = CurBook.Range
         CurRange = CurRange.GoTo(3, 2, 1)
     }
     //插入题型内容
     QuestionTypeNum = QuestionTypeNum + 1
     CurRange.InsertParagraph()
     CurRange.InsertBefore("题型"+String(CurQTypeNum+1))
     //插入bookmark
     l_doc.Bookmarks.Add("QT"+String(CurQTypeNum+1)+"In"+Chaptername, CurRange)
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
     GetNums()
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
//  function InsertQuestion() {
//      let l_doc = wps.WpsApplication().ActiveDocument
//      let bks = l_doc.Bookmarks//获取书签

//      let Chaptername =  document.getElementById('InsertQAChapterSelect').value
//      let Typename = document.getElementById('InsertQAQTSelect').value
//      //未设置题型
//      if(bks.Item(Chaptername).Range.Count == 1) {
//         alert("未设置题型！")
//         return
//     }
//      let MyQuestion = document.getElementById('QInput').value
//      let MyAnswer = document.getElementById('QAInput').value

//      let At = Chapters[Chaptername]//章节标号 
//      //插入bookMark Range内最后一行新段落，并更新bookMark范围，副作用是下一个bookMark Range会改变
     
//      //获得一下题型序号
//     let Type = 1//题型序号
//     for(let i=1; i<=bks.Item(Chaptername).Range.Bookmarks.Count; i++){
//         let BM = bks.Item(Chaptername).Range.Bookmarks.Item(i)
//         console.log(BM.Name)
//         if(BM.Name.indexOf("QT") === 0){
//             if(BM.Name === Typename){
//                 break
//             }else{
//                 Type += 1
//             }
//         }
//     }
     
//      let BNameChapter = "CHP"+String(At)
//      let BNameType = "QT" + String(Type) + "In" + BNameChapter
//      let CurBook = bks.Item(BNameType)
//      let CurRange = CurBook.Range
 
//      //插入题目内容
//      //得到题号
//      let Qnum = 1
//      for(let i=1; i<= CurRange.Bookmarks.Count; i++){
//          if(CurRange.Bookmarks.Item(i).Name.indexOf("D") != -1 && CurRange.Bookmarks.Item(i).Name.indexOf(BNameType) != -1){
//              Qnum = Qnum+1
//          }
//      }
//      let Q = "第"+ String(Qnum) +"题\t\n"//题目标号玄学占位符
//      let A = "第"+ String(Qnum) +"题答案\n"//答案标号
//      CurRange.Paragraphs.Add()
//      CurRange.InsertAfter(Q)
//      //设置题目正文
//      CurRange.Paragraphs.Item(CurRange.Paragraphs.Count-1).Style = 1
//      CurRange.Paragraphs.Item(CurRange.Paragraphs.Count).Style = 1
//      //更新题型bookMark
//      l_doc.Bookmarks.Add(BNameType, CurRange)  
//      //设置题目bookMark*****************
//      let QRange = CurBook.Range
//      //位置需要跳过第*题
//      QRange.Start = CurBook.Range.End+Q.length
//      QRange.End = QRange.Start
//      //选择好区域以插入内容控件
//      QRange.Select()
//      //Question and Answer 两个控件
//      let QControl = InsertTest(Q, MyQuestion)//Question内容控件Range
//      //QControl.Range.InsertParagraph()//插入段落相当于换行
//      QControl.Range.Paragraphs.Add()
//      //Answer 控件的位置要向下移动
//      QRange.Start = QRange.End+MyQuestion.length+3//玄学占位符
//      QRange.End = QRange.Start
//      QRange.Select()
//      let AControl = InsertTest(A, MyAnswer)
//      let MyRange = QControl.Range
//      MyRange.End = AControl.Range.End
//      l_doc.Bookmarks.Add("D"+String(Qnum) + "In"+ BNameType, MyRange)
//      //修正Question内容控件Range
//      //QControl.Range.Text = QControl.Range.Text.slice(0, QControl.Range.Text.length-1)
//      //******************************************************************************* */
     
 
//      //修正下一bookMarkRange
//      let NextBName = "QT" + String(Type+1) + "In" + BNameChapter
//      let NextBookIndex = GetBookMarkIndex(NextBName)
//      if(NextBookIndex === -1){
//          //当前就是最后一个bookmark 无需修改
//          //更新章节书签***********
//          let ChapterBk = bks.Item("CHP"+String(At))
//          let ChapterRange = ChapterBk.Range//定位章节范围
//          let ChNewRange = ChapterRange
//          ChNewRange.End = CurRange.End
//          ChNewRange.Select()
//          l_doc.Bookmarks.Add(Chaptername, ChNewRange)
//          if(ChapterNum >=  At+1){
//              ChNewRange = bks.Item("CHP"+String(At+1)).Range
//              ChNewRange.Start = CurRange.End
//              l_doc.Bookmarks.Add("CHP"+String(At+1), ChNewRange)
//          }
//      }
 
//      //正确的bookmark在上个cur最后一行
//      let NextBook = bks.Item(NextBookIndex)
//      let NextRange = NextBook.Range
//      NextRange.Start = CurRange.End
//      //NextRange.Select()
//      l_doc.Bookmarks.Add(NextBName, NextRange)
 
//  }
function InsertQuestion() {
    let l_doc = wps.WpsApplication().ActiveDocument
    let bks = l_doc.Bookmarks//获取书签

    let Chaptername =  document.getElementById('InsertQAChapterSelect').value
    let Typename = document.getElementById('InsertQAQTSelect').value
    //未设置题型
    if(bks.Item(Chaptername).Range.Count == 1) {
       alert("未设置题型！")
       return
   }
    let MyQuestion = document.getElementById('QInput').value
    let MyAnswer = document.getElementById('QAInput').value

    let At = Chapters[Chaptername]//章节标号 
    //插入bookMark Range内最后一行新段落，并更新bookMark范围，副作用是下一个bookMark Range会改变
    
    //获得一下题型序号
   let Type = 1//题型序号
   for(let i=1; i<=bks.Item(Chaptername).Range.Bookmarks.Count; i++){
       let BM = bks.Item(Chaptername).Range.Bookmarks.Item(i)
       console.log(BM.Name)
       if(BM.Name.indexOf("QT") === 0){
           if(BM.Name === Typename){
               break
           }else{
               Type += 1
           }
       }
   }
    
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
    let Q = "第"+ String(Qnum) +"题: "//题目标号玄学占位符
    let A = "答案： "//答案标号
    // let Range222 = CurRange
    // Range222.Start = CurRange.End+1
    // Range222.End = CurRange.End+1
    // Range222.InsertBefore("\n")
    let CurPara =CurRange.Paragraphs.Add()
    let EndPara = CurRange.Paragraphs.Add()
    let CurStart = CurRange.Start
    let StyleRange = CurRange
    StyleRange.Start = CurRange.Paragraphs.Item(2).Range.Start
    StyleRange.Style = 1
    
    //选择好区域以插入内容控件
    //Question and Answer 两个控件
    let QControl = InsertTest(Q, MyQuestion, CurPara)//Question内容控件Range
    Curpara = QControl.Range.Paragraphs.Add()
    //Answer 控件的位置要向下移动
    let AControl = InsertTest(A, MyAnswer, CurPara)
    Curpara = QControl.Range.Paragraphs.Add()
    let DControl = InsertTest("难度: ", document.getElementById("InsertQD").value, CurPara)
    Curpara = QControl.Range.Paragraphs.Add()
    let TControl = InsertTest("熟练度: ", document.getElementById("InsertQT").value, CurPara)
    
    let MyRange = QControl.Range
    MyRange.End = CurRange.End
    l_doc.Bookmarks.Add("D"+String(Qnum) + "In"+ BNameType, MyRange)

    //更新题型bookmark
    CurRange.Start = CurStart
    CurRange.End = MyRange.End
    l_doc.Bookmarks.Add(BNameType, CurRange) 


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
 function InsertTest(T, Text, CurPara) {
     let l_doc = wps.WpsApplication().ActiveDocument
     let Qtxt = l_doc.ContentControls.Add(wps.wdContentControlText, CurPara.Range);
     Qtxt.Title = T
     Qtxt.MultiLine = true
     Qtxt.Range.Text = T+Text
     Qtxt.Range.Style = 1
     return Qtxt
 }




 // function SetTest(Q) {
 //     let l_doc = wps.WpsApplication().ActiveDocument
 // 	let Qtxt = l_doc.ContentControls.Item(1);
 //     console.log(Qtxt)
 //     Qtxt.Range.Text = "第一题这到底能不能换行换行换行换行换行换行换行换行啊啊啊啊啊啊能不能到底？？？？？"
 
 // }
 
 function test() {
    //  let l_doc = wps.WpsApplication().ActiveDocument
    //  l_doc.Bookmarks.Item("QT3InCHP1").Range.Select()
    //  let se = wps.WpsApplication().Selection
    //  console.log(se.Range.Bookmarks.Count)
    //  for(let i =1;i<=se.Range.Bookmarks.Count;i++){
    //      let a = se.Range.Bookmarks.Item(i)
    //      console.log(a.Name)
    //  }

     //console.log(GetBookMarkIndex("QT3InCHP1"))
     onload()
 }