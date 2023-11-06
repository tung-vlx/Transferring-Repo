document.body.addEventListener('keydown', function(e) {
    if (e.target.id == "segmentNumber" && e.which === 9) {
        let inputTable = document.querySelector('#section1 table.content');
        let newRow = inputTable.insertRow(inputTable.rows.length);
        let segmentItemCell = newRow.insertCell(0);
        segmentItemCell.classList.add("padding-s");
        let segmentItemInput = document.createElement("input");
        segmentItemInput.setAttribute("type","text");
        segmentItemInput.setAttribute("id","segmentItem");
        segmentItemInput.classList.add("padding-s");
        segmentItemCell.appendChild(segmentItemInput);
        let segmentNumberCell = newRow.insertCell(1);
        segmentNumberCell.classList.add("padding-s");
        let segmentNumberInput = document.createElement("input");
        segmentNumberInput.setAttribute("type","text");
        segmentNumberInput.setAttribute("id","segmentNumber");
        segmentNumberInput.classList.add("padding-s");
        segmentNumberCell.appendChild(segmentNumberInput);
    }
    if (e.target.id == "segmentNumber" && e.key === "Enter") {
        segmentDisplay();
    }
})

document.querySelector('#section1 button#submit1').addEventListener('click', function() {
    segmentDisplay();
});

function segmentDisplay() {
    let segmentItemList = document.querySelectorAll("#segmentItem");
    let segmentNumberList = document.querySelectorAll("#segmentNumber");
    let segmentNumberTotal = 0;

    let segmentLength = segmentItemList.length;
    console.log(segmentLength);
    let i = 0;
    let inputTable = document.querySelector('#section1 table.content');
    while (i<segmentLength) {
        if ((segmentItemList[i].value != "") && (segmentNumberList[i].value== "")) {
            segmentNumberList[i].value = "0";
        }
        if ((segmentItemList[i].value == "") && (segmentNumberList[i].value== "")) {
            inputTable.deleteRow(i+1);
            segmentItemList = document.querySelectorAll("#segmentItem");
            segmentNumberList = document.querySelectorAll("#segmentNumber");

            i--;
            segmentLength--;
        }
        i++;
    }
    console.log(segmentItemList.length)
    for (let i = 0; i < segmentNumberList.length; i++) {
        let indexPosition = segmentNumberList[i].value.search(" ");
        if (indexPosition == -1) {
            segmentNumberList[i].value = numClearChar(segmentNumberList[i].value);
            segmentNumberTotal = segmentNumberTotal+parseFloat(segmentNumberList[i].value);                
        } else {
            let numArray = segmentNumberList[i].value.split(" ");
            for (let j = 0; j < numArray.length; j++) {
                if (numArray[j] != "") {
                    segmentNumberTotal = segmentNumberTotal + parseFloat(numClearChar(numArray[j]));
                }
            }
        }
    }

    //segment display --- start
    let output = document.querySelector('#section1 #output1_1');
    output.innerHTML = '';

    let div = document.createElement("div");
    div.classList.add("margin-m");
    let span = document.createElement("span");
    if (segmentNumberList.length == 1) {
        span.innerHTML = "1 segment<br>";
    } else {
        span.innerHTML = segmentNumberList.length + " segments<br>"
    }
    span.classList.add("margin-s")
    div.appendChild(span);
    for (let i = 0; i < segmentItemList.length; i++) {
        let span = document.createElement("span");
        let indexPosition = segmentNumberList[i].value.search(" ");
        if (indexPosition == -1) {
            var percent = parseFloat(segmentNumberList[i].value)/segmentNumberTotal*100;
        } else {
            let total = 0;
            let numArray = segmentNumberList[i].value.split(" ");
            for (let j = 0; j < numArray.length; j++) {
                if (numArray[j] != "") {
                    total = total + parseFloat(numClearChar(numArray[j])); 
                }
            }
            var percent = total/segmentNumberTotal*100;
        }
        span.innerHTML = segmentItemList[i].value + " - " + percent.toFixed(2) + "%<br>";
        span.classList.add("margin-s")
        div.appendChild(span);
    }
    output.appendChild(div);
    let btn = document.createElement("button");
    btn.innerText = "Copy";
    btn.classList.add("tpi-box-orange", "margin-l", "padding-m", "cursor-pointer");
    output.appendChild(btn);
    document.querySelector("#output1_1 button").addEventListener('click', function() {
        copyText("#output1_1 div`");
    })
    //segment display --- end

    //Excel AR calculating bypass --- start
    if (document.querySelector('#section1 input#prefixBln1').checked == true) {
        let excelCell = document.querySelector('#section1 input#prefix1').value;
        let excelColumn = excelCell.replace(/\d/g,'');
        let excelRow = parseFloat(excelCell.replace(/\D/g,''));
        let excelRowLimit = excelRow + segmentNumberList.length - 1;
        output = document.querySelector('div#output1_2');
        output.innerHTML = '';
        let table = document.createElement('table');
        table.style.width = '100%';
        textOutput = '';
        for (let i = 0; i < segmentNumberList.length; i++) {
            let tr = table.insertRow();
            let td = tr.insertCell();
            let indexPosition = segmentNumberList[i].value.search(" ");
            if (indexPosition == -1) {
                textOutput = parseFloat(segmentNumberList[i].value);
            } else {
                textOutput = '=';
                let numArray = segmentNumberList[i].value.split(" ");
                for (let j = 0; j < numArray.length; j++) {
                    if (numArray[j] != "") {
                        if (j==0) {
                            textOutput = textOutput + numClearChar(numArray[j]);                            
                        } else {
                            textOutput = textOutput + "+" + numClearChar(numArray[j]);                        }
                    }
                }
            }

            td.innerText = textOutput;
            td = tr.insertCell();
            let excelRowIndex = excelRow + i;
            td.innerText = "="+excelColumn+excelRowIndex+"/SUM("+excelColumn+excelRow+":"+excelColumn+excelRowLimit+")*100";
            textOutput = 0;
        }
        output.appendChild(table);  
        btn = document.createElement("button");
        btn.innerText = "Copy";
        btn.classList.add("tpi-box-orange", "margin-l", "padding-m", "cursor-pointer");
        output.appendChild(btn);
        document.querySelector("#output1_2 button").addEventListener('click', function() {
            copyText("#output1_2 table");
        })
    }    
    //Excel AR calculating bypass --- end

    resetSection1Input();
}

function resetSection1Input() {
    let inputTable = document.querySelector('#section1 table.content');
    let tableRange = document.querySelectorAll('#section1 table.content tr');
    for (let i = 2; i < tableRange.length; i++) {
        inputTable.deleteRow(2);        
    }
    document.querySelector('#segmentItem').value = '';
    document.querySelector('#segmentNumber').value = '';
}

document.querySelector('input#prefixBln1').addEventListener('click', function() {
    console.log("noted")
    if (document.querySelector('#section1 input#prefixBln1').checked == true) {
        document.querySelector('#section1 input#prefix1').classList.remove("hidden");
    }
    if (document.querySelector('#section1 input#prefixBln1').checked == false) {
        document.querySelector('#section1 input#prefix1').classList.add("hidden");
    }
})


//number format ---> count char
function numClearChar(num) {
    let numTemp = '';
    let countChar = 0;
    let i = 0;
    let char = ",";
    while (countChar < 2 && i < num.length) {
        if (num.charAt(i) == ",") {
            countChar++;
        }  
        i++; 
    }
    if (countChar >= 1) {
        let numTempArray = num.split(',');
        for (let i = 0; i < numTempArray.length; i++) {
            numTemp = numTemp + numTempArray[i];
        }
    }
    if (numTemp != '') {
        num = numTemp;
    }


    char = ".";
    countChar = 0;
    numTemp = '';
    i=0;
    numTempArray =[];
    while (countChar < 2 && i < num.length) {
        if (num.charAt(i) == char) {
            countChar++;
        }  
        i++; 
    }

    if (countChar > 1) {
        numTempArray = num.split('.');
        console.log(numTempArray);
        for (let i = 0; i < numTempArray.length; i++) {
            numTemp = numTemp + numTempArray[i];
        }
    }
    if (numTemp != '') {
        num = numTemp;
    }

    return num;
}
document.querySelector('#section2 input#prefixBln2').addEventListener('click', function() {
    if (document.querySelector('#section2 input#prefixBln2').checked == true) {
        document.querySelector('#section2 input#prefix2').classList.remove("hidden");
    }
    if (document.querySelector('#section2 input#prefixBln2').checked == false) {
        document.querySelector('#section2 input#prefix2').classList.add("hidden");
    }
})

document.querySelector('#section2 textarea#convert2Excel').addEventListener('keypress', function(e) {
    if (e.key === "Enter") {
        convertDisplay();
    }
})
document.querySelector('#section2 button#submit2').addEventListener('click', function() {
    convertDisplay();
})

function convertDisplay() {
    let input = document.querySelector('#section2 textarea#convert2Excel').value;
    input = input.split(/\r?\n|\r|\n/g);
    let output = document.querySelector('#section2 #output2');
    output.innerHTML = '';
    
    let div = document.createElement("div");
    div.classList.add("output_body");
    if (document.querySelector('#section2 input#prefixBln2').checked == true) {
        let prefix = document.querySelector('#section2 input#prefix2').value;
        if (prefix.charAt(prefix.length-1) != " ") {
            prefix = prefix + " ";
        }
        let span = document.createElement("span");
        span.innerText = prefix;
        div.appendChild(span);
    }
    for (let i = 0; i < input.length; i++) {
        if (input[i] && i != (input.length-1)) {
            let span = document.createElement("span");
            let text = input[i].replace("- ","(");
            text = text + "); ";
            span.innerText = text;
            div.appendChild(span);
        }
        if (input[i] && i == (input.length-1)) {
            let span = document.createElement("span");
            let text = input[i].replace("- ","(");
            text = text + ")";
            span.innerText = text;
            div.appendChild(span);
        }
    }
    output.appendChild(div);
    
    let outputFooter = document.createElement("div");
    outputFooter.classList.add("output_footer");
    let btn = document.createElement("button");
    btn.innerText = "Copy";
    btn.classList.add("tpi-box-orange", "margin-l", "padding-m", "cursor-pointer");
    outputFooter.appendChild(btn);
    output.appendChild(outputFooter);
    document.querySelector('#section2 textarea#convert2Excel').value = '';
    document.querySelector('#output2 button').addEventListener('click',function() {
        copyText("#output2 .output_body");
    })
}

function copyText(copyArea) {
    let copyText = document.querySelector(copyArea);
    window.getSelection().selectAllChildren(copyText);
    document.execCommand("Copy");
}
