/**
 * Adminer plugin
 * Download select result as XLSX format.
 * 
 * Install AdminerDumpXlsx to Adminer,
 * and place this file to the plugin directory.
 *
 * Install to Adminer on http://www.adminer.org/plugins/
 * @author Tom Higuchi, http://tom-gs.com/
 */
(function (window, document) {
  /**
   * Create dummy table tag from select result.
   * 
   * @param {HTMLTableElement} tblElem 
   * @param {String} id 
   * @returns {String}
   */
  var createDummyTable = function (tblElem, id) {
    var trs = tblElem.querySelectorAll('tr');
    var html = '<table id="' + id + '" class="table-to-export" data-sheet-name="' + id + '">';
    trs.forEach(function (tr) {
      var ths = tr.querySelectorAll('th');
      var tds = tr.querySelectorAll('td');
      if (ths.length) {
        html += '<tr>';
        ths.forEach(function (th, index) {
          html += '<th>' + getCellValue(th) + '</th>';
        });
        html += '</tr>';
      } else if (tds.length) {
        html += '<tr>';
        tds.forEach(function (td, index) {
          if (skipFirstCell(tblElem, index)) {
            return;
          }
          html += '<td>' + getCellValue(td) + '</td>';
        });
        html += '</tr>';
      }
    });
    html += '</table>';
    return html;
  };

  /**
   * Check if the skippable cell or not.
   * 
   * @param {HTMLTableElement} tblElem 
   * @param {Number} index 
   * @returns {Boolean}
   */
  var skipFirstCell = function (tblElem, index) {
    return 'table' == tblElem.id && 0 === index;
  };

  /**
   * Get plain value in cell. 
   * 
   * @param {HTMLTableElement} cell 
   * @returns {String}
   */
  var getCellValue = function (cell) {
    if ('th' == cell.tagName.toLowerCase()) {
      var ret = cell.id.replace(/th\[(.*)\]/, '$1');
      if (!ret) {
      	var title = cell.hasAttribute('title') ? cell.getAttribute('title') : '';
        ret = title.split('.').slice(-1)[0];
      }
      return ret;
    } else if ('td' == cell.tagName.toLowerCase()) {
      var a = cell.querySelector('a');
      if (a) {
        return a.innerHTML;
      }
      return cell.innerHTML;
    }
  };

  /**
   * Add dump button.
   * 
   * @param {HTMLElement} parent 
   * @param {Number} index 
   */
  var addDumpButton = function (parent, index) {
    var id = 'xlsx-' + index;
    parent.innerHTML += '<button type="button" id="' + id + '" class="button">Download XLSX</button>';
    var dlBtn = document.getElementById(id);
    dlBtn.addEventListener('click', function () {
      dumpXlsx();
    }, false);
  };

  /**
   * Dump table data to XLSX.
   */
  var dumpXlsx = function () {
    // Options for sheets
    var wbopts = {
      bookType: 'xlsx',
      bookSST: false,
      type: 'binary',
      cellText: false,
      cellDates: true
    };

    // Options for workbook
    var wsopts = {
      //header: 1,
      //raw: false,
      dateNF: 'yyyy-mm-dd hh:mm:ss'
    };

    /**
     * Zerofill
     * 
     * @param {Number} number
     * @param {Number} length
     * @returns {String}
     */
    var zerofill = function (number, length) {
      return ('0'.repeat(length) + ('' + number)).slice(-length);
    };

    /**
     * Create file name for download file.
     * 
     * @returns {String}
     */
    var createFileName = function () {
      var fileName = 'adminer.';
      fileName += location.hostname + '.';
      var date = new Date();
      fileName += zerofill(date.getFullYear(), 4);
      fileName += zerofill(date.getMonth() + 1, 2);
      fileName += zerofill(date.getDate(), 2);
      fileName += '_';
      fileName += zerofill(date.getHours(), 2);
      fileName += zerofill(date.getMinutes(), 2);
      fileName += zerofill(date.getSeconds(), 2);
      fileName += '.xlsx';
      return fileName;
    };
  
    var workbook = {SheetNames: [], Sheets: {}};

    document.querySelectorAll('table.table-to-export').forEach(function (currentValue, index) {
      var n = currentValue.getAttribute('data-sheet-name');
      if (!n) {
        n = 'Sheet' + index;
      }
      workbook.SheetNames.push(n);
      workbook.Sheets[n] = XLSX.utils.table_to_sheet(currentValue, wsopts);
    });
  
    var wbout = XLSX.write(workbook, wbopts);
    saveAs(new Blob([s2ab(wbout)], {type: 'application/octet-stream'}), createFileName());
  };

  /**
   * Convert string to ArrayBuffer.
   * 
   * @param {String} s 
   * @returns {ArrayBuffer}
   */
  var s2ab = function (s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) {
      view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
  };

  window.addEventListener('load', function () {
    var div = document.createElement('div');
    div.id = 'dummy-table-area';
    div.style.display = 'none';
    div.style.visibility = 'hidden';
    document.body.appendChild(div);
 
    var table = document.getElementById('table');
    if (table) {
      div.innerHTML += createDummyTable(table, 'table-0');
      let parent = document.querySelector('#fieldset-export .fieldset-content'); // Admin Neo
      if (!parent) {
        parent = document.querySelector('#fieldset-export'); // Adminer
      }
      addDumpButton(parent, 0);
    }

    for (var i = 1; ; i++) {
      var sql = document.getElementById('sql-' + i);
      if (!sql) {
        break;
      }
      var table = sql.nextElementSibling.querySelector('table');
      if (table) {
        div.innerHTML += createDummyTable(table, 'table-' + i);
        let parent = document.querySelector('#export-' + i + ' p'); // Admin Neo
        if (!parent) {
          parent = document.querySelector('#export-' + i); // Adminer
        }
        addDumpButton(parent, i);
      }
    }
  }, false);
})(window, window.document);
