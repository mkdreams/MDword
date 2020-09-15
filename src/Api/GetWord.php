<?php
namespace MDword\Api;
use MDword\Common\Bind;

class GetWord extends Base{
    public function build() {
        $binds = $this->parameters['binds'];
        
        $this->bind($binds);
        return $this->wordProcessor->saveAsContent();
    }
    
    private function bind($binds,$pbind = null,$pBindName = null, $first = true) {
        foreach($binds as $blockName => $value) {
            if($first) {
                $pbind = null;
            }
            
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
                        $this->bind($value['childrens'],$pbind,$blockName,false);
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
        $dataInfo = $data['dataInfo'];
        switch ($dataInfo['type']) {
            case 'local':
                return $this->formatConvert($dataInfo['data'],$data['type']);
                break;
            case 'http':
                if($dataInfo['headers'] === '' && $dataInfo['postData'] === '') {
                    $content = file_get_contents($dataInfo['url']);
                }else{
                    $content = $this->common->CurlSend($dataInfo['url'],$dataInfo['headers'],$dataInfo['postData']);
                }
                return $this->formatConvert($content,$data['type']);
                break;
            case 'excel':
                $data = $this->common->CurlSend($data['url'],$data['headers'],$data['postData']);
                $data = json_decode($data,true);
                return $data;
                break;
        }
    }
    
    private function formatConvert($content,$orgFormat='json') {
        switch ($orgFormat) {
            case 'json':
                return json_decode($content,true);
                break;
            case 'csv':
                return $this->parseCsv($content);
                break;
        }
    }
    
    private function parseCsv($content) {
        $rows = explode("\r", $content);
        foreach($rows as $index => $row) {
            $row = trim($row);
            if($row === '') {
                continue;
            }
            $rows[$index] = explode(',', $row);
        }
        
        $data = [];
        foreach($rows as $index => $row) {
            if($index === 0) {
                continue;
            }
            
            foreach ($row as $key => $value) {
                $data[$index][$key] = $value;
            }
        }
        
        return $data;
    }
}
