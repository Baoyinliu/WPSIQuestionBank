let AK='24.ace996eec431ebfe53e76abba294645e.2592000.1639227418.282335-25154087';

// 新建一个对象，建议只保存一个对象调用服务接口
let client;

function loadOCR(){
    var input = document.getElementById("OCRFile"); //input file
    var file = input.files[0];//获取文件
    document.getElementById('OCRDIV').innerHTML = "";
    var tempHTML = "";
    tempHTML = "<canvas id='OCRCanvas' width='150px' height='100px'></canvas><br>"+
    "<textarea class='TA' type=\"text\" id='OCRtext' cols=\"50\" rows=\"8\" autofocus></textarea>"
    document.getElementById('OCRDIV').innerHTML = tempHTML;
    if (file) {
        //读取本地文件
        var reader = new FileReader();
        var bu = reader.readAsDataURL(file);//ArrayBuffer对象
        console.log(bu)
        reader.onload = function (e) {
            //读取完毕后输出结果
            console.log(e.target.result);
            var url = 'https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic?access_token='+AK;
            var params = {"image":e.target.result}
            $.post(url,
            params,
            function(data,status){
                if(status == 'success'){
                    console.log(data.words_result)
                    var ocr_words = data.words_result;
                    var Stext = '';
                    for(var i=0; i<ocr_words.length; i++){
                        Stext += ocr_words[i].words;
                        Stext += '\n';
                    }
                    document.getElementById('OCRtext').value = Stext; 
                }
            });
            //img show
            var img = new Image();
            img.src = e.target.result;
            img.onload = function(){
                var myCanvas = document.getElementById("OCRCanvas");
                var cxt = myCanvas.getContext('2d');
                console.log(cxt)
                cxt.drawImage(img, 0, 0, 150, 100);
            }
        }
    }

}