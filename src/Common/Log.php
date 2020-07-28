<?php
namespace MDword\Common;

class Log
{
    private $logPath;
    public function __construct() {
        $this->logPath = dirname(__FILE__,3).'/tests/Log/'.date('Ymd').'.log';
        //empty file
        file_put_contents($this->logPath,'');
    }
    
    public function writeLog($content) {
        if(!MDWORD_DEBUG) {
            return ;
        }
        
        $result = file_put_contents($this->logPath, date('H:i').' '.$content."\r\n", FILE_APPEND);
        if($result === false) {
            die('write file error! file:'.$this->logPath);
        }
    }
}