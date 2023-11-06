// Get the input element from the HTML document
var input = document.querySelector("#input");

// Arrays to store comment data
var commentSection = [];
var commentOutput = [];

// Event listener for click events
document.addEventListener("click", function(e) {
    // Display detailed comments
    if (e.target.parentElement.id == "commentDownload" && e.target.nodeName == "BUTTON") {
        displayCommentByFY(e.target.id);
    }

    // Delete and clear old presentation
    if (e.target.parentElement.id == "outputRemove" && e.target.nodeName == "I") {
        removeCommentDisplay();
    }
});

// Event listener for when a file is selected in the input field
input.addEventListener('change', function() {
    checkExcelFile(input.files[0]);
});
// Event listener for when a file is drop in the input field
var dropArea = document.querySelector('#dropArea');

['dragenter', 'dragover', 'dragleave', 'drop'].forEach(e => {
    dropArea.addEventListener(e, function (e) {
        e.preventDefault();
    });
})

dropArea.addEventListener("drop", function handleDrop (e) {
    let data = e.dataTransfer;
    let files = data.files;
    checkExcelFile(files[0]);
});

async function checkExcelFile (files) {
    let rows = [];
    try {
        rows = await readXlsxFile(files);
        readExcelFile(files);
    } catch (error) {
        window.alert(`Move sheet "Exclusion Matrix" to the 1st position`);
        return
    }
}

function readExcelFile (files) {
    // Asynchronously read the uploaded Excel file
    readXlsxFile(files).then(async function(rows) {
        // Determine the start and end of the data table
        let whileBln = true;
        let i = 0;
        let tableStartIndex = 0;
        let tableEndIndex = 0;
        
        // Find the starting row of the data table
        while (i < rows.length && whileBln == true) {
            if (rows[i][0] == "Data Vendor") {
                tableStartIndex = i;
                whileBln = false;
            }
            i++;
        }
        
        // Find the ending row of the data table
        whileBln = true;
        i = tableStartIndex;
        while (i < rows.length && whileBln == true) {
            if (rows[i][0] == null) {
                tableEndIndex = i;
                whileBln = false;
            }
            i++;
        }

        // Identify column indices for Vendor ID and Comments
        let BvDIDSectionIndex = rows[tableStartIndex].indexOf("Vendor ID");
        let commentSectionIndex = rows[tableStartIndex].indexOf("Comments");

        // Extract and process comments
        for (i = tableStartIndex + 1; i < tableEndIndex; i++) {
            commentSection.push([]);
            commentSection[i - tableStartIndex - 1].push([]);
            commentSection[i - tableStartIndex - 1][0] = rows[i][BvDIDSectionIndex];

            if (rows[i][commentSectionIndex] != null) {
                if (rows[i][commentSectionIndex].search(/\nFY|.FY/i) != -1) {
                    let arrayTemp = rows[i][commentSectionIndex].split(/\nFY|.FY/i);
                    let jTemp = 1;
                    for (let j = 0; j < arrayTemp.length; j++) {
                        if (j == 0 && arrayTemp[j] != "") {
                            commentSection[i - tableStartIndex - 1][jTemp] = arrayTemp[j];
                            jTemp++;
                        } else if (arrayTemp[j] != "") {
                            commentSection[i - tableStartIndex - 1][jTemp] = "FY" + arrayTemp[j];
                            jTemp++;
                        }
                    }
                }
            }
        }

        // Clean up the commentSection array
        i = 0;
        while (i < commentSection.length) {
            if (commentSection[i].length == 1) {
                commentSection.splice(i, 1);
            } else {
                i++;
            }
        }

        // Summarize comments into FYxx criteria
        commentToFY();

        // Display comment
        displayComment();

        return commentSection;
    });    
}
// Summarize comments into FYxx criteria
function commentToFY() {
    var commentSectionTemp = commentSection;
    let keyTemp = "";

    for (let i = 0; i < commentSection.length; i++) {
        for (let j = 1; j < commentSection[i].length; j++) {
            let FYIndex = 0;
            let k = 0;
            let whileBln = true;

            // Search for an existing FY criteria in commentOutput
            while (k < commentOutput.length && whileBln == true) {
                if (commentSection[i][j].substring(0, commentSection[i][j].indexOf("-") - 1).replace(/ /g, '') == commentOutput[k][0]) {
                    FYIndex = k;
                    whileBln = false;
                }
                k++;
            }

            if (whileBln == true) {
                // If FY criteria doesn't exist, create a new entry
                commentOutput.push([]);
                commentOutput[commentOutput.length - 1].push([]);
                commentOutput[commentOutput.length - 1][0] = commentSection[i][j].substring(0, commentSection[i][j].indexOf("-") - 1).replace(/ /g, '');
                FYIndex = commentOutput.length - 1;
            }

            commentOutput[FYIndex].push([]);
            // Key: BvD ID
            commentOutput[FYIndex][commentOutput[FYIndex].length - 1].push([]);
            commentOutput[FYIndex][commentOutput[FYIndex].length - 1][0] = commentSection[i][0];

            // Key: FYxx
            commentOutput[FYIndex][commentOutput[FYIndex].length - 1][1] = commentSection[i][j].substring(0, commentSection[i][j].indexOf("-") - 1).replace(/ /g, '');
            commentSectionTemp[i][j] = commentSection[i][j].substring(commentSection[i][j].indexOf("-") + 1);

            // Key: Accept/NP/NF/NI or Non-independent (can be empty "-")
            if (commentSectionTemp[i][j].indexOf("-") == -1) {
                if (commentSectionTemp[i][j].charAt(0) == " ") {
                    commentOutput[FYIndex][commentOutput[FYIndex].length - 1][2] = commentSectionTemp[i][j].substring(1, commentSectionTemp[i][j].length);
                    commentSectionTemp[i][j] = "";
                } else {
                    commentOutput[FYIndex][commentOutput[FYIndex].length - 1][2] = commentSectionTemp[i][j].substring(0, commentSectionTemp[i][j].length);
                    commentSectionTemp[i][j] = "";
                }
            } else {
                commentOutput[FYIndex][commentOutput[FYIndex].length - 1][2] = commentSectionTemp[i][j].substring(0, commentSectionTemp[i][j].indexOf("-") - 1).replace(/ /g, '');
                commentSectionTemp[i][j] = commentSection[i][j].substring(commentSection[i][j].indexOf("-") + 1);
            }

            // Key: MR/W/AR (can be empty)
            if (commentSectionTemp[i][j].substring(0, commentSection[i][j].indexOf("-") - 1).replace(/ /g, '') == "MR" ||
                commentSectionTemp[i][j].substring(0, commentSection[i][j].indexOf("-") - 1).replace(/ /g, '') == "W" ||
                commentSectionTemp[i][j].substring(0, commentSection[i][j].indexOf("-") - 1).replace(/ /g, '') == "AR") {
                commentOutput[FYIndex][commentOutput[FYIndex].length - 1][3] = commentSectionTemp[i][j].substring(0, commentSection[i][j].indexOf("-") - 1).replace(/ /g, '');

                // Key: Comments (can be empty)
                commentSectionTemp[i][j] = commentSection[i][j].substring(commentSection[i][j].indexOf("-") + 1);
                if (commentSectionTemp[i][j].charAt(0) == " ") {
                    commentOutput[FYIndex][commentOutput[FYIndex].length - 1][4] = commentSectionTemp[i][j].substring(1, commentSectionTemp[i][j].length);
                } else {
                    commentOutput[FYIndex][commentOutput[FYIndex].length - 1][4] = commentSectionTemp[i][j].substring(0, commentSectionTemp[i][j].length);
                }
            } else {
                if (commentSectionTemp[i][j].charAt(0) == " ") {
                    commentOutput[FYIndex][commentOutput[FYIndex].length - 1][4] = commentSectionTemp[i][j].substring(1, commentSectionTemp[i][j].length);
                } else {
                    commentOutput[FYIndex][commentOutput[FYIndex].length - 1][4] = commentSectionTemp[i][j].substring(0, commentSectionTemp[i][j].length);
                }
            }
        }
    }
}

// Function to display comment
function displayComment() {
    // Hide the file input container
    document.querySelector('.center-of-screen').classList.add('hidden');

    // Get the output area element
    let outputArea = document.querySelector("#commentDownload");

    // Create buttons for each comment
    for (let i = 0; i < commentOutput.length; i++) {
        let btn = document.createElement('button');
        btn.textContent = commentOutput[i][0];
        btn.id = i;
        btn.classList.add("tpi-button", "margin-m", "cursor-pointer")
        outputArea.appendChild(btn);        
    }

    // Create a remove container
    let div = document.createElement('div');
    div.id = "outputRemove"
    div.innerHTML = '<i class="fa-solid fa-trash"></i>';
    outputArea.appendChild(div);
}

// Function to display detail comment
function displayCommentByFY(commentID) {
    // Toggle active class for buttons
    document.querySelectorAll("#commentDownload button").forEach(function(btn) {
        if (btn.id == commentID) {
            btn.classList.add("active");
        } else {
            btn.classList.remove("active");
        }
    })

    // Get the output area for detail comment
    let outputArea = document.querySelector("#commentPanel");
    outputArea.innerHTML = "";

    // Create a "Copy Table" paragraph with onclick attribute
    let p = document.createElement('p');
    p.innerText = "Copy Table";
    p.style.width = 'fit-content';
    p.classList.add("tpi-button", "margin-m","cursor-pointer")
    p.setAttribute('onclick', `copyText(${commentOutput[commentID][0]})`);
    outputArea.appendChild(p);

    // Create a table for detailed comment
    var table = document.createElement("table");
    table.style.width = '85%'
    table.style.borderCollapse = 'collapse';
    table.style.border = '2px solid var(--text-color)';
    table.style.marginTop = '20px'
    for (let i = 1; i < commentOutput[commentID].length; i++) {
        var tr = table.insertRow(i-1);
        
        for (let j = 0; j < commentOutput[commentID][i].length; j++) {
            var td = tr.insertCell(j);
            td.classList.add("padding-s");
            td.style.border = '1px solid var(--text-color)';
            if (commentOutput[commentID][i][j] == null) {
                td.innerText = "";
            } else if (commentOutput[commentID][i][j].search("\n") == -1) {
                td.innerText = commentOutput[commentID][i][j];
            } else {
                td.innerText = commentOutput[commentID][i][j].replace(/(?:\r\n|\r|\n)/g, '<br>');
            }
        }
    }

    // Set the table ID to the BvD ID
    table.id = commentOutput[commentID][0];
    outputArea.appendChild(table);
}

// Function to copy text to clipboard
function copyText(copyID) {
    let tableId = copyID.id;  // Accessing the 'id' property of the element
    let copyText = document.querySelector(`table#${tableId} tbody`);
    let range = document.createRange();
    range.selectNode(copyText);
    window.getSelection().removeAllRanges(); // Clear any current selection
    window.getSelection().addRange(range); // Add the new selection
    document.execCommand("Copy");
    window.getSelection().removeAllRanges(); // Clear the selection after copying
}

// Function to clear comment display (old session)
function removeCommentDisplay() {
    // Clear the comment download and panel
    document.querySelector('#commentDownload').innerHTML = "";
    document.querySelector('#commentPanel').innerHTML = "";

    // Show the file input container
    document.querySelector('.center-of-screen').classList.remove("hidden");

    // Clear commentSection and commentOutput arrays
    commentSection = [];
    commentOutput = [];

    // Reset the input field
    input.value = "";
}
