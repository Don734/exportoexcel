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
    document.querySelector('.toExcel').addEventListener('click', () => {
        let table = new TableData('tableData');
        table.exportToExcel();
    })
}