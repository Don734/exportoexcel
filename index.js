class TableData {
    constructor(tableID, filename) {
        this.tableID = tableID;
        this.filename = filename ? filename+'.xls' : 'export_data.xls';
        this.dataType = 'data:application/vnd.ms-excel; base64';
        this.template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>';
        this.downloadLink = document.createElement('a');
        this.tableSelect = document.getElementById(tableID);
    }

    base64(s) {
        return window.btoa(unescape(encodeURIComponent(s)));
    }

    format(s, c) {
        return s.replace(/{(\w+)}/g, (m, p) => {
            return c[p];
        })
    }

    exportToExcel() {
        if (!this.tableID.nodeType) {
            let ctx = {worksheet: this.tableID || 'export_data.xls', table: this.tableSelect.innerHTML}
            this.downloadLink.href = 'data:' + this.dataType + ', ' + this.base64(this.format(this.template, ctx));
            this.downloadLink.download = this.filename;
            this.downloadLink.click();
        }
    }
}

document.addEventListener('DOMContentLoaded', Init());

function Init() {
    let span = document.querySelectorAll('span');
    let array = Array.from(span);
    let newArray = sliceIntoChunks(array, 18)
    let newTable = document.createElement('table');

    document.body.appendChild(newTable);
    
    newArray.forEach((trElem, key) => {
        if (key != 33) {
            let tr = document.createElement('tr')
            newTable.appendChild(tr);
            
            trElem.forEach((tdElem) => {
                let td = document.createElement('td');
                tr.appendChild(td);
                console.log(td);
                td.innerText = tdElem.innerText;
            })
        }
    })

    document.querySelectorAll('table td').forEach((elem) => {
        if (elem.innerText == 0 && elem.innerText == '') {
            elem.removeChild(elem);
        }
    })

    console.log(newArray);

    document.querySelector('.toExcel').addEventListener('click', () => {
        let table = new TableData('tableData');
        table.exportToExcel();
    })
}

function sliceIntoChunks(arr, chunkSize) {
    const res = [];
    for (let i = 0; i < arr.length; i+=chunkSize) {
        const chunk = arr.slice(i, i + chunkSize);
        res.push(chunk);
    }
    return res;
}