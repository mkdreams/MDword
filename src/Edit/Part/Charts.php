<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Charts extends PartBase
{
    public $blocks = [];
    public $excel = null;
    public function __construct($word,\DOMDocument $DOMDocument) {
        parent::__construct($word);
        $this->DOMDocument = $DOMDocument;
    }
    
    /**
     *
     * @param string $chartFilenameRel
     * @param [] $datas
     * @param \DOMDocument $domDocument
     * @param string $type
     */
    public function chartRelUpdateByType($datas,$type='str') {
        $refs = $this->DOMDocument->getElementsByTagName($type.'Ref');
        if($refs->length > 0) {
            for($i =0; $i < $refs->length; $i++) {
                $item = $refs->item($i);
                $f = $item->getElementsByTagName('f')->item(0);
                $cache = $item->getElementsByTagName($type.'Cache')->item(0);
                
                foreach($datas as $data) {
                    $range = $f->nodeValue;
                    $rangInfo = $this->parserRange($range);
                    $rangTempInfo = $data[0];
                    $value = $data[1];
                    $action = $data[2];
                    
                    //表不一样，跳过
                    if($rangTempInfo[0] !== $rangInfo[0]) {
                        continue;
                    }
                    //                     var_dump($rangInfo,$rangTempInfo,$value,$action);
                    //赋值
                    if($action == 'set') {
                        if(
                            isset($rangInfo[1][1]) && $rangInfo[1][0] <= $rangTempInfo[1][0] && $rangTempInfo[1][0] <= $rangInfo[1][1]
                            && $rangInfo[2][0] <= $rangTempInfo[2][0] && $rangTempInfo[2][0] <= $rangInfo[2][1]
                            ) {
                                $idx = $rangTempInfo[2][0] - $rangInfo[2][0];
                                $cache->getElementsByTagName('pt')->item($idx)->firstChild->nodeValue = $value;
                            }elseif($rangInfo[1][0] == $rangTempInfo[1][0] && $rangInfo[2][0] == $rangTempInfo[2][0]) {
                                $idx = $rangTempInfo[2][0] - $rangInfo[2][0];
                                $cache->getElementsByTagName('pt')->item($idx)->firstChild->nodeValue = $value;
                            }
                    }elseif($action == 'ext') {//扩展取值范围
                        if(
                            $rangInfo[1][0] === $rangTempInfo[1][0] && $rangInfo[1][1] === $rangTempInfo[1][1]
                            && $rangInfo[2][0] === $rangTempInfo[2][0] && $rangInfo[2][1] === $rangTempInfo[2][1]
                            ) {
                                $f->nodeValue = $value[3];
                                //ptCount update
                                $ptCount = $value[2][1]-$value[2][0]+1;
                                $cache->getElementsByTagName('ptCount')->item(0)->setAttribute('val',$ptCount);
                                
                                $pts = $cache->getElementsByTagName('pt');
                                
                                
                                for($j = $pts->length;$j < $ptCount;$j++) {
                                    $vNode = $this->DOMDocument->createElementNS('http://schemas.openxmlformats.org/drawingml/2006/chart','v','null');
                                    $ptNode = $this->DOMDocument->createElementNS('http://schemas.openxmlformats.org/drawingml/2006/chart','pt');
                                    $ptNode->setAttribute('idx', $j);
                                    $ptNode->appendChild($vNode);
                                    
                                    $cache->appendChild($ptNode);
                                }
                            }
                    }
                }
            }
        }
    }
}
