<?php 
namespace MDword\Convert\Read;

use MDword\Convert\Stream;

class OtherTable{
    private $stream;
    public $versionInfo = [];
    private $types = ["ImageMap"=>0,"ImageMap_Src"=>1,"DocxTheme"=>3];
    /**
     * @param Stream $stream
     */
    public function __construct($stream) {
        $this->stream = $stream;
    }
    
    public function read() {
        $this->stream->enterFrame(4);
        $strLen = $this->stream->getULongLE();
        $this->stream->enterFrame($strLen);
        
        $stCurPos = 0;
        while ($stCurPos < $strLen) {
            $type = $this->stream->getUChar();
            $length = $this->stream->getULongLE();
            
            $this->readOtherContent($type, $length);
//             var_dump($type,$length);exit;
        }
    }
    
    private function readOtherContent($type,$length) {
        switch ($type) {
            case $this->types['ImageMap']:
                break;
            case $this->types['ImageMap_Src']:
                break;
            case $this->types['DocxTheme']:
                $this->readTheme();
                break;
        }
    }
    
    private function readTheme() {
        
    }
}