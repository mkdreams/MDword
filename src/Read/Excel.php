<?php
namespace MDword\Read;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Excel
{
    /**
     * @var Spreadsheet
     */
    private $spreadsheet = null;
    public function __construct() {
        $this->spreadsheet = new Spreadsheet();
    }
    
    public function load($archive) {
        $this->tempDocumentFilename = tempnam($this->getTempDir(), 'MDword');
        if (false === $this->tempDocumentFilename) {
            throw new \Exception('temp path make faild!');
        }
        
        if (false === copy($archive, $this->tempDocumentFilename)) {
            throw new \Exception($archive.'copy file fiald!');
        }
        
        
        $this->zip = new \ZipArchive();
        $this->zip->open($this->tempDocumentFilename);
        
        $this->read();
    }
}
