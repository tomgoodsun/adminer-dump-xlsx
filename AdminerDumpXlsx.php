<?php
/**
 * Displays column comment in table header data-listing page
 *
 * Install to Adminer on http://www.adminer.org/plugins/
 * @author Tom Higuchi, http://tom-gs.com/
 */

class AdminerDumpXlsx
{
    private $pathToSheetJs;
    private $pathToFileSaverJs;
    private $pathtoDumpXlsxJs;

    /**
     * @param string $pathToSheetJs
     * @param string $pathToFileSaverJs
     * @param string $pathtoDumpXlsxJs
     */
    public function __construct(
        $pathToSheetJs = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.9.10/xlsx.full.min.js',
        $pathToFileSaverJs = 'https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/1.3.3/FileSaver.min.js',
        $pathtoDumpXlsxJs = 'dumpxlsx.js'
    ) {
        $this->pathToSheetJs = $pathToSheetJs;
        $this->pathToFileSaverJs = $pathToFileSaverJs;
        $this->pathToDumpXlsxJs = $pathtoDumpXlsxJs;
    }

    public function head()
    {
        echo script_src($this->pathToSheetJs);
        echo script_src($this->pathToFileSaverJs);
        echo script_src($this->pathToDumpXlsxJs);
    }
}
