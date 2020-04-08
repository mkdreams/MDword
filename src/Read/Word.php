<?php
namespace MDword\Read;

use MDword\Edit\Part\Document;
use MDword\Edit\Part\Comments;
use MDword\Read\Part\ContentTypes;

class Word 
{
    public $zip = null;
    
    private $Content_Types = null;
    
    private $tempDocumentFilename = null;
    /**
     * @var Document
     */
    public $documentEdit = null;
    /**
     * @var Comments
     */
    public $commentsEdit = null;
    
    public $parts = [];
    
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
    
    public function save()
    {
        $this->deleteComments();
        
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
        
        if(MDWORD_DEBUG === true) {
            $this->zip->open($this->tempDocumentFilename);
            $this->zip->extractTo(MDWORD_GENERATED_DIRECTORY);
            $this->zip->close();
        }
        
        return $this->tempDocumentFilename;
    }
    
    
    private function deleteComments() {
        $parts = [];
        if(isset($this->parts[15])) {
            $parts = array_merge($parts,$this->parts[15]);
            unset($this->parts[15]);
        }
        if(isset($this->parts[16])) {
            $parts = array_merge($parts,$this->parts[16]);
            unset($this->parts[16]);
        }
        if(isset($this->parts[17])) {
            $parts = array_merge($parts,$this->parts[17]);
            unset($this->parts[17]);
        }
        
        foreach($parts as $part) {
            $this->zip->deleteName($part['PartName']);
        }
        
        //remove marked
        $this->documentEdit->deleteMarked();
        
        //test
//         echo $this->documentEdit->DOMDocument->saveXML();exit;
        
        //remove comments tag
        $DOMDocument = $this->documentEdit->DOMDocument;
        $commentRangeStarts = $DOMDocument->getElementsByTagName('commentRangeStart');
        $commentRangeEnds = $DOMDocument->getElementsByTagName('commentRangeEnd');
        $commentReferences = $DOMDocument->getElementsByTagName('commentReference');
        
        while ($commentRangeStart = $commentRangeStarts->item(0)) {
            $commentRangeStart->parentNode->removeChild($commentRangeStart);
        }
        while ($commentRangeEnd = $commentRangeEnds->item(0)) {
            $commentRangeEnd->parentNode->removeChild($commentRangeEnd);
        }
        while ($commentReference = $commentReferences->item(0)) {
            $commentReference->parentNode->parentNode->removeChild($commentReference->parentNode);
        }
    }
    
    /**
     *
     * @param string $filename
     * @return \DOMDocument
     */
    public function getXmlDom($filename) {
        $xml = $this->zip->getFromName($filename);
        $domDocument = new \DOMDocument();
        $domDocument->loadXML($xml);
        if(MDWORD_DEBUG === true) {
            $domDocument->formatOutput = true;
        }
        
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
