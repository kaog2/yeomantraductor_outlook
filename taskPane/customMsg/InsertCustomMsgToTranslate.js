// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
///Puefo dejar la app como ejemplo para futuras notificaciones en un taskPane
/// <reference path="../App.js" />
(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
      $(document).ready(function () {
          $('#insertText').click(insertText);
      });
  };
  
  function insertText(event) {
    var textToInsert = $('#textToInsert').val();
    

    // Insert as plain text (CoercionType.Text)
    Office.context.mailbox.item.body.setSelectedDataAsync(
      textToInsert, 
      { coercionType: Office.CoercionType.Text });
  }

})();
/*(function () {
  "use strict";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
      $(document).ready(function () {
          app.initialize();
          $('#insertText').click(insertText);
      });
  };
  
  function insertText(event) {
    var textToInsert = $('#textToInsert').val();
    

    // Insert as plain text (CoercionType.Text)
    Office.context.mailbox.item.body.setSelectedDataAsync(
      textToInsert, 
      { coercionType: Office.CoercionType.Text }, 
      function (asyncResult) {
        // Display the result to the user
        if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
          app.showNotification("Success", "\"" + textToInsert + "\" inserted successfully.");
        }
        else {
          app.showNotification("Error", "Failed to insert \"" + textToInsert + "\": " + asyncResult.error.message);
        }
      });
  }

})();*/