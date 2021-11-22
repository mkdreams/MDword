<?php
namespace MDword\Common;

use MDword\WordProcessor;

class Bind
{
    private $data;
    /**
     * @var WordProcessor
     */
    private $wordProcessor;
    
    private $pre = '';
    
    public function __construct($wordProcessor,$data,$pre='') {
        $this->wordProcessor = $wordProcessor;
        $this->data = $data;
        $this->pre = $pre;
    }
    
    public function bindValue($name,$keyList,$pBindName=null,$callbackOrValueType=null,$emptyCallBack=null) {
        static $binds = [];
        
        //loop
        if(!is_null($pBindName) && isset($binds[$pBindName])) {
            foreach($binds[$pBindName] as $bind) {
                $bind->bindValue($name,$keyList,null,$callbackOrValueType,$emptyCallBack);
            }
            
            return $this;
        }
        
        if('INNER_VARS' !== $pBindName) {
            $data = $this->data;
        }else{
            $data = $this->wordProcessor->getInnerVars();
        }
        foreach($keyList as $key) {
            $data = $data[$key];
        }

        if(is_array($data)) {
            $count = count($data);
            $this->wordProcessor->clones($name.$this->pre,$count);
            $i = 0;
            foreach($data as $subData) {
                if(!isset($binds[$name])) {
                    $binds[$name] = [];
                }
                $binds[$name][] = new Bind($this->wordProcessor, $subData, $this->pre.'#'.$i++);
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