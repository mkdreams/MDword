<?php
namespace MDword\Api;
use MDword\Common\Bind;

class GetWord extends Base{
    public function build() {
        $binds = $this->parameters['binds'];
        
        $this->bind($binds);
        return $this->wordProcessor->saveAsContent();
    }
    
    private function bind($binds,$pbind = null,$pBindName = null) {
        foreach($binds as $blockName => $value) {
            if(is_array($value)) {
                if(isset($value['dataKeyName'])) {
                    /**
                     * @var Bind $pbind
                     */
                    $pbind = $this->wordProcessor->getBind($this->getData($value['dataKeyName']));
                }
                
                if(is_null($pbind)) {
                    $this->wordProcessor->setValue($blockName, $value);
                }else{
                    $pbind->bindValue($blockName, $value['keyList'], $pBindName);
                    if(isset($value['childrens'])) {
                        $this->bind($value['childrens'],$pbind,$blockName);
                    }
                }
            }else{
                $this->wordProcessor->setValue($blockName, $value);
            }
        }
    }
    
    private function getData($name) {
        static $datas = [];
        if(isset($datas[$name])) {
            return $datas[$name];
        }
        
        if(isset($this->parameters['datas'][$name])) {
            return $this->parseData($this->parameters['datas'][$name]);
        }
        
        return null;
    }
    
    
    private function parseData($data) {
        switch ($data['type']) {
            case 'json':
                return $data['data'];
                break;
        }
    }
}
