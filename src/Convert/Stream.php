<?php
namespace MDword\Convert;

class Stream {
    private $pos = 0;//end pos
    private $cur = 0;//start pos
    private $size = 0;
    private $data = [];
    public function __construct($szSrc, $streamType = 1) {
        if($streamType === 1) {//base64
            $strs = explode(';', $szSrc, 2);
            $this->size = intval($strs[0]);
            $chars = base64_decode($strs[1]);
            unset($szSrc);
        }elseif($streamType === 2) {//bin
            $chars = $szSrc;
            unset($szSrc);
        }
        
        $chartsLen = strlen($chars);
        for($i = 0; $i < $chartsLen; $i++) {
            $this->data[] = ord($chars[$i]);
        }
        $this->size = count($this->data);
        var_dump($this->data);
    }
    
    public function getVersionInfo() {
        $semicolonCount = 0;
        $info = [];
        for($i = 0; $i < $this->size; $i++) {
            if($this->data[$i] === 59) {//59 is ord(';')
                $semicolonCount++;
            }else{
                $info[$semicolonCount] .= chr($this->data[$i]);
            }
            
            if($semicolonCount === 3) {
                break;
            }
        }
        
        $this->pos += $i + 1;//\n
        
        return $info;
    }
    
    public function enterFrame($count=1) {
        $this->cur = $this->pos;
        $this->pos += $count;
        var_dump($this->cur,$this->pos);
    }
    
    public function getUChar() {
        return $this->data[$this->cur++];
    }
    
    public function getULongLE() {
        return $this->getLong();
    }
    
    public function getChar() {
        $m = $this->data[$this->cur++];
        $m = $m>127?$m-256:$m;
        return $m;
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
}

