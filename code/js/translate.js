let AK='24.be09c7d0bb7930ca874d51b9d25aa3a1.2592000.1639393820.282335-25163857';
let PreSe = '';

setInterval(function(){
    var se = wps.WpsApplication().Selection
    //console.log(se.Text)
    //console.log(PreSe)
    if(se.Text != PreSe){
        loadTrans();
        PreSe = se.Text;
    }
},1000);


function loadTrans(){
    var se = wps.WpsApplication().Selection
    var tempHTML = "";
    for(var i=0; i<se.Text.length; i++){
        if(i%60 == 0 && i>0)
            tempHTML += "<br>";
        tempHTML += se.Text[i];
    }
    document.getElementById('TransSrc').innerHTML = tempHTML;

    var params = {"q":"hello","from":"en","to":"zh"}
    var src = se.Text;
    var p = '&q='+src+'&from=en&to=zh';
    var url = 'https://aip.baidubce.com/rpc/2.0/mt/texttrans/v1?access_token='+AK+p;
    $.post(url,
    params,
    function(data,status){
        console.log(data)
        if('result' in data){
            var tempHTML2 = "";
            var DstText = data.result.trans_result[0].dst;
            console.log(DstText)
            for(var i=0; i<DstText.length; i++){
                if(i%30 == 0 && i>0)
                    tempHTML2 += "<br>";
                tempHTML2 += DstText[i];
            }
            console.log(tempHTML2)
            document.getElementById('TransDst').innerHTML = tempHTML2;
        }
    });
           
}