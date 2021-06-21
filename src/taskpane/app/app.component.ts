import { Component, Input } from "@angular/core";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
const template = require("./app.component.html");
/* global require, Word */

@Component({
  selector: "app-home",
  templateUrl: "./src/taskpane/app/app.component.html",
  styleUrls: ["./src/taskpane/app/app.component.css"],
})
export class AppComponent {
  visibleName: string = "";
  uniqueName: string = "";
  phText: string = "";
  presetSelected: string = "";
  presets: IPreset[] = [
    {
      visibleName: "Arranger Name",
      uniqueName: "arrangerName",
      phText: "Enter Arranger Name Here",
    },
    {
      visibleName: "Coupon Type",
      uniqueName: "couponType",
      phText: "Enter Coupon Type Here",
    },
    {
      visibleName: "Issuer Name",
      uniqueName: "issuerName",
      phText: "Enter Issuer Name Here",
    },
    {
      visibleName: "Issuer Region",
      uniqueName: "issuerRegion",
      phText: "Enter Issuer Region Here",
    },
    {
      visibleName: "Program Currency",
      uniqueName: "programCurrency",
      phText: "Enter Program Currency Here",
    },
    {
      visibleName: "Program Amount",
      uniqueName: "programAmount",
      phText: "Enter Program Amount Here",
    },
    {
      visibleName: "Tenor",
      uniqueName: "tenor",
      phText: "Enter Tenor Here",
    },
    {
      visibleName: "Product Type",
      uniqueName: "productType",
      phText: "Enter Product Type Here",
    },
    {
      visibleName: "Product Type Short",
      uniqueName: "productTypeShort",
      phText: "Enter Product Type Short Here",
    },
    {
      visibleName: "Issuer Rating",
      uniqueName: "issuerRating",
      phText: "Enter Issuer Rating Here",
    },
    {
      visibleName: "Issuer Rating Org",
      uniqueName: "issuerRatingOrg",
      phText: "Enter Issuer Rating Org Here",
    },
    {
      visibleName: "Principal Amount Currency",
      uniqueName: "principalAmountCurrency",
      phText: "Enter Principal Amount Currency Here",
    },
    {
      visibleName: "Principal Amount",
      uniqueName: "principalAmount",
      phText: "Enter Principal Amount Here",
    },
    {
      visibleName: "Pricing Date",
      uniqueName: "pricingDate",
      phText: "Enter Pricing Date Here",
    },
    {
      visibleName: "Deposit Date",
      uniqueName: "depositDate",
      phText: "Enter Deposit Date Here",
    },
  ];

  constructor() {}
  ngOnInit() {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Word) {
        // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
        if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
          console.log("Sorry. The add-in uses Word.js APIs that are not available in your version of Office.");
        }
      }
    });
  }

  componentWillMount() {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Word) {
        // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
        if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
          console.log("Sorry. The add-in uses Word.js APIs that are not available in your version of Office.");
        }
      }
    });
  }

  insertSampleData() {
    Word.run(function (context) {
      var docBody = context.document.body;
      docBody.insertParagraph(
        "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
        "Start"
      );
      return context.sync();
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  setPresetValues() {
    let result = this.presets.find(({ visibleName }) => visibleName === this.presetSelected);
    this.visibleName = result.visibleName;
    this.uniqueName = result.uniqueName;
    this.phText = result.phText;
  }

  createContentControl() {
    Word.run((context) => {
      // Queue commands to create a content control.
      var serviceNameRange = context.document.getSelection();
      var serviceNameContentControl = serviceNameRange.insertContentControl();
      //get values from text box to set as property of content control
      serviceNameContentControl.title = this.visibleName;
      serviceNameContentControl.tag = this.uniqueName;
      serviceNameContentControl.placeholderText = this.phText;
      serviceNameContentControl.appearance = "Tags";
      serviceNameContentControl.color = "blue";

      return context.sync();
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  clearAnnotationFields() {
    this.visibleName = "";
    this.uniqueName = "";
    this.phText = "";
  }

  clearContentInControl() {
    Word.run(function (context) {
      var myCCs = context.document.getSelection().contentControls;
      context.load(myCCs);
      return context.sync().then(function () {
        for (var i = 0; i < myCCs.items.length; i++) {
          // here you will get the full content of content controls within the selection,
          var serviceNameContentControl = myCCs.items[i];
          serviceNameContentControl.insertText("", "Replace");
          console.log("this is full  paragraph:" + (i + 1) + ":" + myCCs.items[i].text);
        }
      });
    }).catch(function (error) {
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
      return context.sync().then(function () {
        var div = document.getElementById("exportedFields");
        div.innerHTML = "";
        for (var i = 0; i < myCCs.items.length; i++) {
          div.innerHTML += myCCs.items[i].tag;
          div.innerHTML += " : ";
          div.innerHTML += myCCs.items[i].text;
          div.innerHTML += "</br>";

          // here you will get the full content of content controls within the selection,
          console.log("this is full  paragraph:" + (i + 1) + ":" + myCCs.items[i].text);
        }
      });
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  loadContentInControl() {
    Word.run(function (context) {
      //for now this set static data for all fields, but this can be modified to
      // fetch data from server based on tags and replace the respective values.
      var myCCs = context.document.getSelection().contentControls;
      context.load(myCCs);
      return context.sync().then(function () {
        for (var i = 0; i < myCCs.items.length; i++) {
          // here you will get the full content of content controls within the selection,
          var serviceNameContentControl = myCCs.items[i];
          serviceNameContentControl.insertText("some data", "Replace");
          console.log("this is full  paragraph:" + (i + 1) + ":" + myCCs.items[i].text);
        }
      });
    }).catch(function (error) {
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
      return context.sync().then(function () {
        var obj = {};
        for (var i = 0; i < myCCs.items.length; i++) {
          var key = myCCs.items[i].tag;
          var value = myCCs.items[i].text;
          obj[key] = value;
        }
        var json = JSON.stringify(obj, null, 4);
        // this can be sent to a server instead of just displaying
        var div = document.getElementById("exportedFields");
        div.innerHTML = json;
      });
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }
}

interface IPreset {
  visibleName: string;
  uniqueName: string;
  phText: string;
}

//   async run() {
//     return Word.run(async (context) => {
//       /**
//        * Insert your Word code here
//        */

//       // insert a paragraph at the end of the document.
//       const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

//       // change the paragraph color to blue.
//       paragraph.font.color = "blue";

//       await context.sync();
//     });
//   }
// }
