// Select the table, rows input, columns input, and initialize tableExists flag
let table = document.getElementsByClassName("sheet-body")[0],
rows = document.getElementsByClassName("rows")[0],
columns = document.getElementsByClassName("columns")[0]
tableExists = false

// Function to generate the table
const generateTable = () => {
    let rowsNumber = parseInt(rows.value), columnsNumber = parseInt(columns.value)
    table.innerHTML = ""
  if (rowsNumber > 0 && columnsNumber > 0) {
    // Generate the table if the rows and columns are valid
    for (let i = 0; i < rowsNumber; i++) {
      let tableRow = "";
      for (let j = 0; j < columnsNumber; j++) {
        tableRow += `<td contenteditable></td>`;
      }
      table.innerHTML += tableRow;
    }
    tableExists = true;
  } else {
    // Display SweetAlert.js alert when fields are empty
    Swal.fire("Empty Fields", "Please enter the number of rows and columns.", "warning");
  }
};

// Function to export the table
const ExportToExcel = (type, fn, dl) => {
     // Display SweetAlert.js alert when no table is generated
    if(!tableExists){
        Swal.fire("No Table", "Please generate a table before exporting.", "warning");
        return
    }
    var elt = table
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
};
