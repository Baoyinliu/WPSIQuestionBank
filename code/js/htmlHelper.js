

function SaveAsOFD(){
    console.log("111")
    console.log(wps.WpsApplication().ActiveDocument.SaveAs2("111.ofd",  102, false, undefined, true, undefined, false, false, false, false, false, 0, false, false, 0, false));
}

 function InsertQASelectChp(){
    GetNums()
    let Chaptername =  document.getElementById('InsertQAChapterSelect').value
    let At = Chapters[Chaptername]//章节标号
    let CurQTypeNum = 0//题目类型数量
    let l_doc = wps.WpsApplication().ActiveDocument
    let bks = l_doc.Bookmarks//获取书签
    let ChapterBk = bks.Item(Chaptername)
    let ChapterRange = ChapterBk.Range//定位章节范围
     //获得一下题型数量
    for(let i=1; i<=ChapterRange.Bookmarks.Count; i++){
        let BM = ChapterRange.Bookmarks.Item(i)
        if(BM.Name.indexOf("QT") === 0){
            //console.log(BM.Name)
            CurQTypeNum = CurQTypeNum + 1
        }
    }


    document.getElementById('InsertQASelectQT').innerHTML = ""
    let tempHTML ="<select id=\"InsertQAQTSelect\">"
    let flag = 0
    for(let i=1; i<=CurQTypeNum; i++){
        let CurType = "QT"+String(i)+"In"+Chaptername
        tempHTML = tempHTML + "<option value = "+ CurType + ">"+QTypes[CurType]+"</option>"
        flag = 1
    }
    if(flag == 1){
        tempHTML = tempHTML + "<span>请先选择章节</span>"
    }
    tempHTML = tempHTML + "</select>"
    document.getElementById('InsertQASelectQT').innerHTML = tempHTML+ "<button class='btn' onclick = \"GetNums()\">确定</button class='btn'>"
}

function QInputType(){
    if(document.getElementById('QInputType').value == "text"){
        document.getElementById('InsertQuestion').innerHTML = ""
        let QHTML =  "<select id=\"QInputType\">" +"<option value = \"text\">文本</option> "
        +"<option value = \"img\">图片</option>" 
        +"</select>"
        QHTML += " <button class='btn' onclick = \"GetNums()\">确定</button class='btn'> "
        QHTML += "<textarea class='TA' type=\"text\" id=\"QInput\" cols=\"30\" rows=\"4\" autofocus></textarea>"
        document.getElementById('InsertQuestion').innerHTML = QHTML 
    }else{
        document.getElementById('InsertQuestion').innerHTML = ""
        let QHTML =  "<select id=\"QInputType\">"
        +"<option value = \"img\">图片</option>" 
        +"<option value = \"text\">文本</option> "
        +"</select>"
        QHTML += " <button class='btn' onclick = \"GetNums()\">确定</button class='btn'> "
        QHTML += "<input type=\"image\" id=\"QInput\" height=\"80\", width=\"80\">"
        document.getElementById('InsertQuestion').innerHTML = QHTML 
    } 
    
 }

 function QAInputType(){
    if(document.getElementById('QAInputType').value == "text"){
        document.getElementById('InsertQAnswer').innerHTML = ""
        let QHTML =  "<select id=\"QAInputType\">" +"<option value = \"text\">文本</option> "
        +"<option value = \"img\">图片</option>" 
        +"</select>"
        QHTML += " <button class='btn' onclick = \"GetNums()\">确定</button class='btn'> "
        QHTML += "<textarea  class='TA' type=\"text\" id=\"QAInput\" cols=\"30\" rows=\"4\" autofocus></textarea>"
        document.getElementById('InsertQAnswer').innerHTML = QHTML 
    }else{
        document.getElementById('InsertQAnswer').innerHTML = ""
        let QHTML =  "<select id=\"QAInputType\">"
        +"<option value = \"img\">图片</option>" 
        +"<option value = \"text\">文本</option> "
        +"</select>"
        QHTML += " <button class='btn' onclick = \"GetNums()\">确定</button class='btn'> "
        QHTML += "<input type=\"image\" id=\"QAInput\" height=\"80\", width=\"80\">"
        document.getElementById('InsertQAnswer').innerHTML = QHTML 
    } 
    document.getElementById('InsertQA').innerHTML = " <button class='ant-btn ant-btn-red' onclick = \"InsertQuestion()\">插入题目</button class='btn'> "
    
 }