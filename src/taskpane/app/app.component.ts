import { Component } from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
const template = require("./app.component.html");
/* global require, Word */

@Component({
  selector: "app-home",
  template,
})
export default class AppComponent {
  welcomeMessage = "Welcome";
  contractNumber = "";
  issueDate: Date = null;

  insertSampleData() {
    Word.run(function (context) {
        var docBody = context.document.body;
        docBody.insertParagraph("Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
            "Start");

        return context.sync();
    })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
  }

  createContentControl() {
    Word.run(function (context) {

        // Queue commands to create a content control.
        var serviceNameRange = context.document.getSelection();
        var serviceNameContentControl = serviceNameRange.insertContentControl();
        //get values from text box to set as property of content control
        serviceNameContentControl.title = (<HTMLInputElement>document.getElementById("vname")).value;
        serviceNameContentControl.tag = (<HTMLInputElement>document.getElementById("uname")).value;
        serviceNameContentControl.appearance = "Tags";
        serviceNameContentControl.color = "blue";

        return context.sync();
    })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
}

clearContentInControl() {
    Word.run(function (context) {
        var myCCs = context.document.getSelection().contentControls;
        context.load(myCCs);
        return context.sync()
            .then(function () {
                for (var i = 0; i < myCCs.items.length; i++) {
                    // here you will get the full content of content controls within the selection,
                    var serviceNameContentControl = myCCs.items[i]
                    serviceNameContentControl.insertText("", "Replace");
                    console.log("this is full  paragraph:" + (i + 1) + ":" + myCCs.items[i].text);

                }
            })
    })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
}

showContentInControl() {
    Word.run(function (context) {
        var myCCs = context.document.getSelection().contentControls;
        context.load(myCCs);
        return context.sync()
            .then(function () {
                var div = document.getElementById('extractedFields');
                div.innerHTML = ''
                for (var i = 0; i < myCCs.items.length; i++) {

                    div.innerHTML += myCCs.items[i].tag;
                    div.innerHTML += ' : ';
                    div.innerHTML += myCCs.items[i].text;
                    div.innerHTML += '</br>';

                    // here you will get the full content of content controls within the selection,
                    console.log("this is full  paragraph:" + (i + 1) + ":" + myCCs.items[i].text);

                }
            })
    })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
}


loadData() {
    Word.run(function (context) {
        //for now this set static data for all fields, but this can be modified to
        // fetch data from server based on tags and replace the respective values.
        var myCCs = context.document.getSelection().contentControls;
        context.load(myCCs);
        return context.sync()
            .then(function () {
                for (var i = 0; i < myCCs.items.length; i++) {
                    // here you will get the full content of content controls within the selection,
                    var serviceNameContentControl = myCCs.items[i]
                    serviceNameContentControl.insertText("some data", "Replace");
                    console.log("this is full  paragraph:" + (i + 1) + ":" + myCCs.items[i].text);

                }
            })
    })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
}

exportContentInControl() {
    Word.run(function (context) {
        var myCCs = context.document.getSelection().contentControls;
        context.load(myCCs);
        return context.sync()
            .then(function () {
                var obj = {};
                for (var i = 0; i < myCCs.items.length; i++) {
                    var key = myCCs.items[i].tag;
                    var value = myCCs.items[i].text;
                    obj[key] = value;
                }
                var json = JSON.stringify(obj);
                // this can be sent to a server instead of just displaying
                var div = document.getElementById('exportedFields');
                div.innerHTML = json
            })
    })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
}

  async run() {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

      // change the paragraph color to blue.
      paragraph.font.color = "blue";

      await context.sync();
    });
  }
}
