<?php
namespace MDword\Common;

use MDword\WordProcessor;

class Bind
{
    private $data = [];
    /**
     * @var WordProcessor
     */
    private $wordProcessor = null;

    private $binds = [];
    
    private $pre = '';
    
    public function __construct($wordProcessor,$data,$pre='') {
        if(!is_null($wordProcessor)) {
            $this->wordProcessor = $wordProcessor;
        }
        $this->data = $data;

        $this->pre = $pre;
    }

    public function bindValue($name,$keyList,$pBindName=null,$callbackOrValueType=null,$emptyCallBack=null) {
        //loop
        if(!is_null($pBindName) && isset($this->binds[$pBindName])) {
            foreach($this->binds[$pBindName] as $bind) {
                $bind->bindValue($name,$keyList,null,$callbackOrValueType,$emptyCallBack);
            }
            
            return $this;
        }
        
        $data = $this->data;
        foreach($keyList as $key) {
            $data = $data[$key];
        }
        
        if(is_array($data)) {
            $count = count($data);
            $this->wordProcessor->clones($name.$this->pre,$count);
            $i = 0;
            foreach($data as $subData) {
                if(!isset($this->binds[$name])) {
                    $this->binds[$name] = [];
                }
                $this->binds[$name][] = new Bind($this->wordProcessor, $subData, $this->pre.'#'.$i++);
            }
            
            if($count === 0 && !is_null($emptyCallBack)) {
                $this->wordProcessor->cloneTo($name.$this->pre,$emptyCallBack($data,$this->data),$count);
            }
        }else{
            $type = MDWORD_TEXT;
            if(!is_null($callbackOrValueType)) {
                if(is_callable($callbackOrValueType)) {
                    $data = $callbackOrValueType($data,$this->data,$this->pre);
                    if(isset($data['overrideType'])) {
                        $type = $data['overrideType'];
                        $data = $data['overrideValue'];
                    }
                }elseif(is_int($callbackOrValueType)) {
                    $type = $callbackOrValueType;
                }
            }
            $this->wordProcessor->setValue($name.$this->pre,$data,$type);
        }
        
        return $this;
    }
}