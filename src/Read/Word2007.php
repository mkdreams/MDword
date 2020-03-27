<?php
namespace PhpOffice\PhpWord\Processor\Read;

use PhpOffice\PhpWord\Processor\Read\Part\ContentTypes;

class Word2007 
{
    private $zip = null;
    
    private $Content_Types = null;
    
    public function load($archive) {
        $this->zip = new \ZipArchive();
        $this->zip->open($archive);
        $this->read();
    }
    
    private function read() {
        $this->Content_Types = new ContentTypes($this->getXmlDom('[Content_Types].xml'));
        
        $files = $this->getZipFiles();
    }
    
    /**
     *
     * @param string $filename
     * @return \DOMDocument
     */
    private function getXmlDom($filename) {
        $xml = $this->zip->getFromName($filename);
        $domDocument = new \DOMDocument();
        $domDocument->loadXML($xml);
        
        return $domDocument;
    }
    
    protected function getZipFiles() {
        static $pathIndx = null;
        
        if(is_null($pathIndx)) {
            for($i=0;$i < $this->zip->numFiles; $i++) {
                $name = $this->zip->getNameIndex($i);
                $nameArr = explode('/', $name);
                
                $item = &$pathIndx;
                foreach ($nameArr as $pathOrName) {
                    if(!isset($item[$pathOrName])) {
                        $item[$pathOrName] = [];
                    }
                    $item = &$item[$pathOrName];
                }
            }
            unset($item);
        }
        
        return $pathIndx;
    }
    
}
