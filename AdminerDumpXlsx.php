<?php
/**
 * Adminer plugin
 * Download select result as XLSX format.
 *
 * Install to Adminer on http://www.adminer.org/plugins/
 * @author Tom Higuchi, http://tom-gs.com/
 */

class AdminerDumpXlsx
{
    /**
     * @var  string
     */
    private $pathToSheetJs;

    /**
     * @var string
     */
    private $pathToFileSaverJs;

    /**
     * @var string
     */
    private $pathToDumpXlsxJs;

    /**
     * @param string|null $pathToSheetJs
     * @param string|null $pathToFileSaverJs
     * @param string|null $pathtoDumpXlsxJs
     */
    public function __construct($pathToSheetJs = null, $pathToFileSaverJs = null, $pathToDumpXlsxJs = null)
    {
        $this->pathToSheetJs = $pathToSheetJs ?? 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.16.9/xlsx.full.min.js';
        $this->pathToFileSaverJs = $pathToFileSaverJs ?? 'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js';

        $dumpxlsxjs = /** @lang javascript */
        <<<'JS'
        !function(e,t){let n=null,l=function(){return null===n&&(n="zeroadmin"),n},r=function(e,t){let n=e.querySelectorAll("tr"),l=`<table id="${t}" class="table-to-export" data-sheet-name="${t}">`;return n.forEach(function(t,n){let r="td",a="td",u=[];0===n&&(r="th",a="th, td"),u=t.querySelectorAll(a),u.length&&(l+="<tr>",u.forEach(function(t,n){o(e,n)||(l+=`<${r}>`+i(t)+`</${r}>`)}),l+="</tr>")}),l+="</table>",l},o=function(e,t){return"table"==e.id&&0===t},i=function(e){if("th"==e.tagName.toLowerCase()){let t=e.id.replace(/th\[(.*)\]/,"$1");if(!t){t=(e.hasAttribute("title")?e.getAttribute("title"):"").split(".").slice(-1)[0]}return t}if("td"==e.tagName.toLowerCase()){let t=e.querySelector("a");return t?t.innerHTML:e.innerHTML}},a=function(e,n){let r="xlsx-"+n,o="adminer"==l()?"&nbsp;":"";e.innerHTML+=`${o}<button type="button" id="${r}" class="button">Download XLSX</button>`,t.getElementById(r).addEventListener("click",function(){c()},!1)},u=function(e,t){return("0".repeat(t)+""+e).slice(-t)},c=function(){let e={dateNF:"yyyy-mm-dd hh:mm:ss"},n={SheetNames:[],Sheets:{}};t.querySelectorAll("table.table-to-export").forEach(function(t,l){let r=t.getAttribute("data-sheet-name");r||(r="Sheet"+l),n.SheetNames.push(r),n.Sheets[r]=XLSX.utils.table_to_sheet(t,e)});let r=XLSX.write(n,{bookType:"xlsx",bookSST:!1,type:"binary",cellText:!1,cellDates:!0});saveAs(new Blob([s(r)],{type:"application/octet-stream"}),function(){let e=l()+".";e+=location.hostname+".";let t=new Date;return e+=u(t.getFullYear(),4),e+=u(t.getMonth()+1,2),e+=u(t.getDate(),2),e+="_",e+=u(t.getHours(),2),e+=u(t.getMinutes(),2),e+=u(t.getSeconds(),2),e+=".xlsx",e}())},s=function(e){let t=new ArrayBuffer(e.length),n=new Uint8Array(t);for(let t=0;t!=e.length;++t)n[t]=255&e.charCodeAt(t);return t},d=function(e,n){let l=t.querySelector(e);return l||(l=t.querySelector(n)),l};e.addEventListener("load",function(){let e=t.createElement("div");e.id="dummy-table-area",e.style.display="none",e.style.visibility="hidden",t.body.appendChild(e),function(e){let n=t.getElementById("table");if(n){e.innerHTML+=r(n,"table-0");let t=d("#fieldset-export .fieldset-content","#fieldset-export");a(t,0)}}(e),function(e){for(let n=1;;n++){let l=t.getElementById(`sql-${n}`);if(!l)break;let o=l.nextElementSibling.querySelector("table");if(o){e.innerHTML+=r(o,`table-${n}`);let t=d(`#export-${n} p`,`#export-${n}`);a(t,n)}}}(e)},!1)}(window,window.document);
        JS;
        printf("<script nonce='%s'>%s</script>\n", Adminer\get_nonce(), $dumpxlsxjs);

    }

    /**
     * @param string $path
     * @return string
     */
    private function addNonce($path)
    {
        if (strpos('?', $path)) {
            return $path .= '&nonce=' . Adminer\get_nonce();
        }
        return $path .= '?nonce=' . Adminer\get_nonce();
    }

    public function head()
    {
        echo Adminer\script_src($this->addNonce($this->pathToSheetJs));
        echo Adminer\script_src($this->addNonce($this->pathToFileSaverJs));
    }
}
