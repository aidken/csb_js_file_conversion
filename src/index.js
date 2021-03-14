import "./styles.css";
import XLSX from "xlsx";
import papa from "papaparse";

// cSpell:ignore xlsx papaparse codesandbox noopener noreferrer

let orders = {};

let Order = function(row) {
  this.itemNumber = row.B.toString();
  this.qty        = row.E;
};

let inventories = {};

let Inventory = function (record) {
  this.itemNumber  = record[0].toString();
  this.warehouse   = record[1];
  this.location    = record[2].toString();
  this.qty         = record[3];
  this.tenDigitLot = record[4].toString();
};

let ExcelToJSON = function() {

  this.parseExcel = function(file) {
    var reader = new FileReader();

    reader.onload = function(evt) {
      let data          = evt.target.result;
      let workbook      = XLSX.read(data, {type: 'binary'});
      let worksheet     = workbook.Sheets['worksheet you want to read'];
      let XL_row_object = XLSX.utils.sheet_to_json(worksheet, {header: "A"});

      XL_row_object.forEach(function(row) {
        // write parsing process here
        if (typeof row.B !== 'undefined' && row.B.toString() !== '商品ｺｰﾄﾞ' && row.B.toString() !== '総計') {
          let order = new Order(row);
          if (order.itemNumber in orders) {
            console.log(`Strange, this item number ${order.itemNumber} appears more than once.`);
            orders[order.itemNumber] += order.qty;
          } else {
            orders[order.itemNumber] = order.qty;
          }
        }  // end if
      }); // end XL_row_object.forEach

    };  // end onload

    reader.onerror = function(ex) {
      console.log(ex);
    };  // end reader.onerror

    reader.readAsBinaryString(file);

  }; // end parseExcel

}; // end function ExcelToJSON

function handleFileSelect1(evt) {
  let files = evt.target.files; // FileList object
  let xl2json = new ExcelToJSON();
  xl2json.parseExcel(files[0]);
}; // end function handleFileSelectOrders


function handleFileSelect2(evt) {

  let file = evt.target.files[0];
  papa.parse(file, {
    header: false,
    dynamicTyping: true,
    // your parsing process goes here
    complete: function (results) {
      results.data.forEach(function (d) {
        if (d.length === 5) {
          let x = new Inventory(d);
          if (x.warehouse === "FX") {
            if (x.itemNumber in inventories) {
              inventories[x.itemNumber] += x.qty;
            } else {
              inventories[x.itemNumber] = x.qty;
            }
          }
        }
      }); // end forEach

    },
  });
} // end handleFileSelectInventories

// initiate app
document.getElementById("app").innerHTML = `
<h2 class="jumbotron text-center" style="margin-bottom:0">Javascript Text/Excel File Uploader</h2>

<div class="container" style="margin-top:20px">

  <div class="row">

    <div class="col-sm-3">
      <ul>
        <li><a href="https://codesandbox.io/" target="_blank" rel="noopener noreferrer">Codesandbox</a></li>
        <li><a href="https://github.com/" target="_blank" rel="noopener noreferrer">Github</a></li>
        <li><a href="https://stackoverflow.com/" target="_blank" rel="noopener noreferrer">Stack Overflow</a></li>
      </ul>
    </div>

    <div class="col-sm-9">

      <h4>How to Use</h4>
      <p>This is a <a href="https://codesandbox.io/">Codesandbox</a> template of
      text/Excel file uploader. This lets you upload file and do something to
      the files.</p>

      <form enctype="multipart/form-data">
        <p>
          <label for='upload1'>Upload Excel File</label>
          <input id="upload1" type=file name="files1[]" accept='.xlsm, .xlsx'>
        </p>
      </form>

      <form enctype="multipart/form-data">
        <p>
          <label for='upload2'>Upload Text File</label>
          <input id="upload2" type=file name="files2[]" accept='.txt, .csv'>
        </p>
      </form>

      <h4>History</h4>
      <ul>
        <li>March 14th 2020: Created.</li>
      </ul>

    </div>

  </div>

</div>
`;

document
  .getElementById("upload1")
  .addEventListener("change", handleFileSelect1, false);

document
  .getElementById("upload2")
  .addEventListener("change", handleFileSelect2, false);
