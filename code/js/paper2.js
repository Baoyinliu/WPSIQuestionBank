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
 function onload2() {
    if (typeof (wps.Enum) != "object") { // 如果没有内置枚举值
        wps.Enum = WPS_Enum
    }
    GetNums()
    GetPaperTitle()
    GetPaper()
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
       let QD = wps.PluginStorage.getItem("QD")
       let QT = wps.PluginStorage.getItem("QT")
       
       InsertTest(CHPI, ChpQTypeNum, QNum, QD, QT)
   }
}

//生成试题
function InsertTest(CHPI, ChpQTypeNum, QNum, QD, QT){

   let cnt=0

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
                   cnt = cnt+1
                   //插入HTML(题目)
                   console.log(QuestionText)
                   PreHTML = document.getElementById('PaperQuesDiv2').innerHTML//保存前面的html语句
                   innerHtml = ""
                   innerHtml = PreHTML +"<div class= \"PaperQues\" > <textarea class='TA' cols=50 rows=5>"+ QuestionText +"</textarea>" 
                   +"<textarea class='TA' cols=50 rows=5>"+ "填写答案：" +"</textarea>" 
                   +"<textarea class='TA' id='A_" + cnt  + "' style='display:none' cols=50 rows=5>"+ AText +"</textarea>" 
                   +"<br><button class='ant-btn ant-btn-red'onclick=\"CorrectPaper(" + cnt + ")\")>" + "显示答案" + "</button><br><br> "
                   +
                   "</div>"
                   document.getElementById('PaperQuesDiv2').innerHTML = innerHtml
               }
           }
       }
      
       //console.log(ChpBks.Item(TypeBookName).Range)
       console.log("dsadasdasd")
   }

   wps.PluginStorage.setItem("CNT", cnt)

}

function CorrectPaper(AIndex){
    var traget=document.getElementById("A_"+AIndex);  
    if(traget.style.display=="none"){  
        traget.style.display="";  
    }else{  
        traget.style.display="none";  
    }  

}

function CorrectPaper2(){
    for(let AIndex=1; AIndex<=wps.PluginStorage.getItem("CNT") ; AIndex++){
        
        var traget=document.getElementById("A_"+AIndex);  
        if(traget.style.display=="none"){  
            traget.style.display="";  
        }else{  
            traget.style.display="none";  
        }  
    }

}
