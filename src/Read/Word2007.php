<?php
namespace MDword\Read;

use MDword\Read\Part\ContentTypes;

class Word2007 
{
    private $zip = null;
    
    private $Content_Types = null;
    
    private $tempDocumentFilename = null;
    
    public $parts = [];
    
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
    
    public static function getTempDir()
    {
        $tempDir = sys_get_temp_dir();
        
        if (!empty(self::$tempDir)) {
            $tempDir = self::$tempDir;
        }
        
        return $tempDir;
    }
    
    
    private function read() {
        $this->Content_Types = new ContentTypes($this->getXmlDom('[Content_Types].xml'));
        foreach ($this->Content_Types->overrides as $part) {
            if($part['ContentType'] === 14) {//image/png
                $this->parts[$part['ContentType']][] = ['PartName'=>$part['PartName'],'DOMElement'=>$part['PartName']];
            }else{
                $this->parts[$part['ContentType']][] = ['PartName'=>$part['PartName'],'DOMElement'=>$this->getXmlDom($part['PartName'])];
            }
        }
    }
    
    public function save()
    {
        foreach($this->parts as $list ) {
            foreach($list as $part) {
                if(is_object($part['DOMElement'])) {
                    $this->zip->addFromString($part['PartName'], $part['DOMElement']->saveXML());
                }
            }
        }
        
        if (false === $this->zip->close()) {
            throw new \Exception('Could not close zip file.');
        }
        
        return $this->tempDocumentFilename;
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
