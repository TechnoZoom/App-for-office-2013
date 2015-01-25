/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }
})();

function writeData() {
    Office.context.document.setSelectedDataAsync("∫", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}
function writeToPage(text) {
    document.getElementById('results').innerText = text;
}

function writeData_2() {
    Office.context.document.setSelectedDataAsync("∞", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}

function writeData_3() {
    Office.context.document.setSelectedDataAsync("≠", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



  
function writeData_4() {
    Office.context.document.setSelectedDataAsync("±", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



  
function writeData_5() {
    Office.context.document.setSelectedDataAsync("∓", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


 

function writeData_6() {
    Office.context.document.setSelectedDataAsync("×", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


  

function writeData_7() {
    Office.context.document.setSelectedDataAsync("⊗", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


   

function writeData_8() {
    Office.context.document.setSelectedDataAsync("÷", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



  
function writeData_9() {
    Office.context.document.setSelectedDataAsync("≈", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    

function writeData_10() {
    Office.context.document.setSelectedDataAsync("⇒", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_11() {
    Office.context.document.setSelectedDataAsync("→", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    
function writeData_12() {
    Office.context.document.setSelectedDataAsync("⇔", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_13() {
    Office.context.document.setSelectedDataAsync("∀", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


 


function writeData_14() {
    Office.context.document.setSelectedDataAsync("≅", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



      
function writeData_15() {
    Office.context.document.setSelectedDataAsync("~", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_16() {
    Office.context.document.setSelectedDataAsync("δ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_17() {
    Office.context.document.setSelectedDataAsync("∮", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    

function writeData_18() {
    Office.context.document.setSelectedDataAsync("π", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_19() {
    Office.context.document.setSelectedDataAsync("σ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    
function writeData_20() {
    Office.context.document.setSelectedDataAsync("⊥", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_21() {
    Office.context.document.setSelectedDataAsync("∬", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



    
function writeData_22() {
    Office.context.document.setSelectedDataAsync("∭", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_23() {
    Office.context.document.setSelectedDataAsync("∯", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


      


function writeData_24() {
    Office.context.document.setSelectedDataAsync("∰", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_25() {
    Office.context.document.setSelectedDataAsync("∑", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_26() {
    Office.context.document.setSelectedDataAsync("∴", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    

function writeData_27() {
    Office.context.document.setSelectedDataAsync("∵", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_28() {
    Office.context.document.setSelectedDataAsync("∈", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_29() {
    Office.context.document.setSelectedDataAsync("∉", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}













        
function writeData_p() {
    Office.context.document.setSelectedDataAsync("α", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_p_1() {
    Office.context.document.setSelectedDataAsync("β", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_p_2() {
    Office.context.document.setSelectedDataAsync("γ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_p_3() {
    Office.context.document.setSelectedDataAsync("δ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    


function writeData_p_4() {
    Office.context.document.setSelectedDataAsync("ε", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_p_5() {
    Office.context.document.setSelectedDataAsync("ζ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    

function writeData_p_6() {
    Office.context.document.setSelectedDataAsync("η", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_p_7() {
    Office.context.document.setSelectedDataAsync("θ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


   

function writeData_p_8() {
    Office.context.document.setSelectedDataAsync("κ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_p_9() {
    Office.context.document.setSelectedDataAsync("λ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


      

function writeData_p_10() {
    Office.context.document.setSelectedDataAsync("μ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_p_11() {
    Office.context.document.setSelectedDataAsync("ν", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_p_12() {
    Office.context.document.setSelectedDataAsync("ω", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


  

function writeData_p_13() {
    Office.context.document.setSelectedDataAsync("ψ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



  
function writeData_p_14() {
    Office.context.document.setSelectedDataAsync("χ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


 

function writeData_p_15() {
    Office.context.document.setSelectedDataAsync("φ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_p_16() {
    Office.context.document.setSelectedDataAsync("τ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}







          

function writeData_m_1() {
    Office.context.document.setSelectedDataAsync("♩", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_m_2() {
    Office.context.document.setSelectedDataAsync("♪", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    
function writeData_m_3() {
    Office.context.document.setSelectedDataAsync("♫", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    
function writeData_m_4() {
    Office.context.document.setSelectedDataAsync("♬", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}

  

function writeData_m_5() {
    Office.context.document.setSelectedDataAsync("♭", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    
function writeData_m_6() {
    Office.context.document.setSelectedDataAsync("♮", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}





function writeData_m_7() {
    Office.context.document.setSelectedDataAsync("♯", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}








           

function writeData_e_1() {
    Office.context.document.setSelectedDataAsync("≧◠‿●‿◠≦", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_2() {
    Office.context.document.setSelectedDataAsync("(ô‿ô)", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_3() {
    Office.context.document.setSelectedDataAsync("٩(●̮̃•)۶", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}

                            

function writeData_e_4() {
    Office.context.document.setSelectedDataAsync("(っ◔◡◔)っ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_5() {
    Office.context.document.setSelectedDataAsync("♥", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_6() {
    Office.context.document.setSelectedDataAsync("ʃ(˘▽ƪ)", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_7() {
    Office.context.document.setSelectedDataAsync(" ⁀⊙﹏☉⁀ ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}

                  


function writeData_e_8() {
    Office.context.document.setSelectedDataAsync("(⊙.⊙(☉_☉)⊙.⊙)", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_9() {
    Office.context.document.setSelectedDataAsync("(◑_◑)", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_10() {
    Office.context.document.setSelectedDataAsync("ಥ_ಥ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}

         

function writeData_e_11() {
    Office.context.document.setSelectedDataAsync("(╥﹏╥)", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_12() {
    Office.context.document.setSelectedDataAsync(" :-Þ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_13() {
    Office.context.document.setSelectedDataAsync(":-C", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}

         

function writeData_e_14() {
    Office.context.document.setSelectedDataAsync("✿◕ ‿ ◕✿", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_15() {
    Office.context.document.setSelectedDataAsync("(^,^)", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}

                   

function writeData_e_16() {
    Office.context.document.setSelectedDataAsync("O.o", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_17() {
    Office.context.document.setSelectedDataAsync("(。_。)", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_18() {
    Office.context.document.setSelectedDataAsync("【・ヘ・?】", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}

              

function writeData_e_19() {
    Office.context.document.setSelectedDataAsync("~(˘▾˘)~", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_20() {
    Office.context.document.setSelectedDataAsync("(✖╭╮✖)", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}

         

function writeData_e_21() {
    Office.context.document.setSelectedDataAsync("( ►_◄ )", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_22() {
    Office.context.document.setSelectedDataAsync("(ﾉ◕ヮ◕)ﾉ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


         
function writeData_e_23() {
    Office.context.document.setSelectedDataAsync("t(-_-t)", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}



function writeData_e_24() {
    Office.context.document.setSelectedDataAsync("┬┴┬┴┤(･_├┬┴┬┴", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}






      


function writeData_po_1() {
    Office.context.document.setSelectedDataAsync("ʳ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_2() {
    Office.context.document.setSelectedDataAsync("ɹ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_3() {
    Office.context.document.setSelectedDataAsync("ɾ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    
function writeData_po_4() {
    Office.context.document.setSelectedDataAsync("ʃ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_5() {
    Office.context.document.setSelectedDataAsync("θ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    


function writeData_po_6() {
    Office.context.document.setSelectedDataAsync("t̬", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_7() {
    Office.context.document.setSelectedDataAsync("ʊ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    

function writeData_po_8() {
    Office.context.document.setSelectedDataAsync("ʊ̈", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_9() {
    Office.context.document.setSelectedDataAsync("ʌ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


      
function writeData_po_10() {
    Office.context.document.setSelectedDataAsync("ʒ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_11() {
    Office.context.document.setSelectedDataAsync("ʔ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_12() {
    Office.context.document.setSelectedDataAsync("æ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


      


function writeData_po_13() {
    Office.context.document.setSelectedDataAsync("ɑ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_14() {
    Office.context.document.setSelectedDataAsync("ð", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_15() {
    Office.context.document.setSelectedDataAsync("ə", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


    

function writeData_po_16() {
    Office.context.document.setSelectedDataAsync("ɚ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_17() {
    Office.context.document.setSelectedDataAsync("ɜ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


  

function writeData_po_18() {
    Office.context.document.setSelectedDataAsync("ɛ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_19() {
    Office.context.document.setSelectedDataAsync("ɝ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_20() {
    Office.context.document.setSelectedDataAsync("ɪ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


      

function writeData_po_21() {
    Office.context.document.setSelectedDataAsync("ɪ̈", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_22() {
    Office.context.document.setSelectedDataAsync("ɫ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_23() {
    Office.context.document.setSelectedDataAsync("ŋ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}


  

function writeData_po_24() {
    Office.context.document.setSelectedDataAsync("ɔ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}




function writeData_po_25() {
    Office.context.document.setSelectedDataAsync("ɒ", function (asyncResult) {
        if (asyncResult.status === "failed") {
            writeToPage('Error: ' + asyncResult.error.message);
        }
    });
}







