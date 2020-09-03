<?php
use MDword\Convert\Stream;
use MDword\Convert\WordParse;

require_once(__DIR__.'/../../Autoloader.php');
require_once(__DIR__.'/../../src/Config/Main.php');

$stream = new Stream(file_get_contents(MDWORD_SRC_DIRECTORY.'/../tests/samples/convert/Editor.bin'),2);

$WordParse = new WordParse($stream);
$WordParse->readBin();
var_dump($WordParse->versionInfo);exit;
/*
$parse = new parse('96;AgAAADEA//8BAGzHPq2l+QEApwAAAAEAAAAAAAAAAAAAAAAAAAAAAAAA9v///y4AAABAAEAAVgBlAHIAcwBpAG8AbgAuAEAAQABCAHUAaQBsAGQALgBAAEAAUgBlAHYA',0);

$classId = $parse->getString();
$nChangesType = $parse->getLong();
var_dump($classId,$nChangesType);

$parse->tableInfos();

class parse {
    private $cur = 0;
    private $size = 0;
    private $data = [];
    public function __construct($szSrc, $offset) {
        $strs = explode(';', $szSrc, 2);
        $this->size = intval($strs[0]);
        $chars = base64_decode($strs[1]);
        $chartsLen = strlen($chars);
        for($i = 0; $i < $chartsLen; $i++) {
            $this->data[] = ord($chars[$i]);
        }
        var_dump($this->data);
    }
    
    public function getLong() {
        return $this->data[$this->cur++] | $this->data[$this->cur++]<<8 | $this->data[$this->cur++]<<16 | $this->data[$this->cur++]<<24;
    }
    
    public function getString() {
        $len = $this->getLong();
        
        $a = '';
        for($i = 0; $i < $len; $i = $i + 2) {
            $a .= chr($this->data[$this->cur + i] | $this->data[$this->cur + i + 1]<<8);
        }
        
        $this->cur += $len;
        
        return $a;
    }
    
    public function tableInfos() {
        $FileCheckSum  = $this->getLong();
        $FileSize      = $this->getLong();
        $Description   = $this->getLong();
        $ItemsCount    = $this->getLong();
        $PointIndex    = $this->getLong();
        $StartPoint    = $this->getLong();
        $LastPoint     = $this->getLong();
        $SumIndex      = $this->getLong();
        $DeletedIndex  = $this->getLong();
        $VersionString = $this->getString();
        
        var_dump($FileCheckSum,$FileSize,$Description,$ItemsCount,$PointIndex,$StartPoint,$LastPoint,$SumIndex,$DeletedIndex,$VersionString);
    }
    
}
*/
