var wordFilePath = "";
var excelFileContent = "";

/* Load docx file to the browser */
function loadFile(url,callback){
  JSZipUtils.getBinaryContent(url,callback);
}

$(document).ready (function () {
  $('#uploadForm').submit(function() {
    $('#input_excel').css('display', 'none');
    $('#inputExcelUploadStatus').text("File is uploading...");
    $(this).ajaxSubmit({
      error: function(xhr) {
        status('Error: ' + xhr.status);
        $('#input_excel').css('display', 'inline');
        excelFileContent = "";
      },

      success: function(res) {
        $('#input_excel').css('display', 'inline');
        $('#inputExcelUploadStatus').text("");
        excelFileContent = new Object();
        var text = "[" + '\n';
        for (var i = 0; i < res.data.length; i ++) {
          var eachObj = res.data[i];
          text += "   {" + '\n';
          text += "       name :    " + eachObj.name + '\n';
          text += "       value :    " + eachObj.value + '\n';
          text += "   }" + '\n';

          excelFileContent[eachObj.name] = eachObj.value;
        }
        text += "]";
        $('#input_excel_data').text(text);

        console.log(excelFileContent);

        alert('Uploaded successfully!');
      }
    });
    return false;
  });

  $('#uploadFormWord').submit(function() {
    $('#input_word').css('display', 'none');
    $('#inputWordUploadStatus').text("File is uploading...");
    $(this).ajaxSubmit({
      error: function(xhr) {
        //status('Error: ' + xhr.status);
        //status('Error: ', xhr);
        console.log('Error Status: ' + xhr);
        $('#input_word').css('display', 'inline');
        wordFilePath = "";
      },

      success: function(res) {
        $('#input_word').css('display', 'inline');
        $('#inputWordUploadStatus').text("");

        wordFilePath = res.data;

        alert('Uploaded successfully!');
      }
    });
    return false;
  });

  $('#downloadWord').on('click', function() {
    if (wordFilePath != "" && excelFileContent != "") {
      loadFile(wordFilePath, function(error,content){
        if (error) { throw error };
        var zip = new JSZip(content);
        var doc = new Docxtemplater().loadZip(zip);
        /*doc.setData({
          first_name: 'John',
          last_name: 'Doe',
          phone: '0652455478',
          description: 'New Website'
        });*/
        doc.setData(excelFileContent);
        try {
          // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
          doc.render();
        } catch (error) {
          var e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
          }
          console.log(JSON.stringify({error: e}));
          // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
          throw error;
        }

        var out=doc.getZip().generate({
          type:"blob",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }) //Output the document using Data-URI
        saveAs(out,"output.docx")
      });
    } else {
      alert("Please input source files!!!");
    }
  });
});
