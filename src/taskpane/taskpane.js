/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
export var old_text = "";
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
  }
});

// export async function run() {
//   return Word.run(async (context) => {
//     /**
//      * Insert your Word code here
//      */
//     // get content of cursor active paragraph


//     setInterval(async  function () {

//       var paragraphx = context.document.getSelection();
//       context.load(paragraphx);
//       await context.sync();


//       // debugger
//       // if paragraphx.text is empty, then get the previous paragraph
//       // setinterval
//       // check if different from old_text
      
//       if (paragraphx.text !== "" && paragraphx.text !== old_text) {
//         old_text = paragraphx.text;
//         document.getElementById("cursorHighlight").innerHTML = paragraphx.text;
//         console.log(paragraphx.text);
//         // const word = old_text
//         // url = 'http://localhost:8000/api/Question/?format=json&search=' + paragraphx.text
//         console.log('http://localhost:8000/api/Question/?format=json&search='+old_text)
//         fetch('http://localhost:8000/api/Question/?format=json&search='+old_text)
//                 .then(response => response.json())
//                 .then(response => {
//                   console.log(JSON.stringify(response))
//                   $("#display_question").html('')
//                   console.log(JSON.stringify(response))
//                   response = response['results']
//                   for (const key in response) {
//                       if (response.hasOwnProperty(key)) {
//                           $("#display_question").append(
//                             '<div class="row">' +
//                             '<div class="col-sm-10 alert alert-primary" role="alert" id=${key}>'
//                                 +
//                               // ' '+'<button onclick=clone(this)>'+
//                             ' ' + '<i class="fas fa-clone" onclick="clone(this, \'' + response[key].name + '\')"></i>&nbsp;&nbsp;'+
//                                '<div class="hide_overflow">'+
//                                 response[key].name
//                                 +  '</div>'
//                             +  '</div>'
//                             +
//                             '<div class="alert alert-secondary" role="alert">'  +
//                             response[key].date + '</div>'
//                             + '</div>'
//                         )

//                         // const btn = document.getElementById("btn_clone");
//                         // btn.addEventListener("click", clone)



                       
//                           }
//                         }


//                         // var showChar = 200;
//                         // var ellipsestext = "...";
//                         // var moretext = "Read more";
//                         // var lesstext = "Read less";
                       
//                         // $('.hide_overflow').each(function() {
//                         //    var content = $(this).html();
//                         //    if(content.length > showChar) {
//                         //      c = content.substr(0, showChar);
//                         //     var h = content.substr(showChar, content.length - showChar);
//                         //     var html = c + '<span class="moreelipses">' + ellipsestext + '</span><a href="" class="morelink">'+ moretext + '</a></span>';
//                         //     $(this).html(html);
//                         //   }

                          
//                         // });


//                         // $(document).on("click", ".morelink", function() {
//                         //   if($(this).hasClass("less")) {
//                         //     $(this).removeClass("less");
//                         //     $(this).html(moretext);
//                         //   } else {
//                         //     $(this).addClass("less");
//                         //     // var html = c + '<span class="moreelipses">' + ellipsestext + '</span><span class="morecontent"><span>' + h + '</span><a href="" class="morelink">' + lesstext + '</a></span>'
//                         //     $(this).html(lesstext);
//                         //   }
//                         //   $(this).parent().prev().toggle();
//                         //   $(this).prev().toggle();
//                         //   return false;
//                         // });
                       










//                 })
                
 
//       }

     
//     }, 1000);
    
    

//     // function clone(event) {
//     //   Word.run(async (context) => {
//     //     // Get the selected paragraph
//     //     const paragraphx = context.document.getSelection();
//     //     context.load(paragraphx);
//     //     await context.sync();
    
//     //     // Get the text content of the parent element of the clicked button
//     //     const value = event.target.parentElement.textContent;
    
//     //     // Insert the text into the selected paragraph, replacing its contents
//     //     paragraphx.insertText(value, Word.InsertLocation.replace);
    
//     //     document.getElementById("run").style.display = "none";
//     //     await context.sync();
//     //   });
//     // }


//       // carry out the function here
  



//     // insert a paragraph at the end of the document.
//     // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

//     document.getElementById("run").style.display = "none";
//     await context.sync();
//   });
// }

// function clone(event) {
//   Word.run(async (context) => {
//     // Get the selected paragraph
//     const paragraphx = context.document.getSelection();
//     context.load(paragraphx);
//     await context.sync();
//     console.log("clone called")

//     // Get the text content of the parent element of the clicked button
//     const value = event.target.parentElement.textContent;

//     // Insert the text into the selected paragraph, replacing its contents
//     paragraphx.insertText(value, Word.InsertLocation.replace);

//     await context.sync();
//   });
// }