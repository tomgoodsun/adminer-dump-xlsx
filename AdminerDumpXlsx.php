<?php
/**
 * Adminer plugin
 * Download select result to XLSX format.
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
        $this->pathToSheetJs = $pathToSheetJs ?? 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.9.10/xlsx.full.min.js';
        $this->pathToFileSaverJs = $pathToFileSaverJs ?? 'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/1.3.3/FileSaver.min.js';
        $this->pathToDumpXlsxJs = $pathToDumpXlsxJs ?? 'dumpxlsx.js';
    }

    public function head()
    {
        echo script_src($this->pathToSheetJs);
        echo script_src($this->pathToFileSaverJs);
        echo script_src($this->pathToDumpXlsxJs);
    }
}
