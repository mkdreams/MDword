<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;
use MDword\XmlTemple\XmlFromPhpword;

class Numbering extends PartBase
{
    private $abstractNums = null;
    private $numIdToAbstractNumId = [];
    public function __construct($word,\DOMDocument $DOMDocument) {
        parent::__construct($word);
        
        $this->DOMDocument = $DOMDocument;
        $this->initNameSpaces();
    }

    public function getAbstractNumById($numId) {
        if(is_null($this->abstractNums)) {
            $tempnums = $this->DOMDocument->getElementsByTagName('abstractNum');
            $nums = $this->DOMDocument->getElementsByTagName('num');
            foreach($nums as $num) {
                $abstractNumId = intval($this->getVal($num,'abstractNumId'));
                $this->numIdToAbstractNumId[intval($this->getAttr($num,'numId'))] = [$abstractNumId,$num];
            }
            $this->abstractNums = [];
            foreach($tempnums as $abstractNum) {
                $abstractNumId = $this->getAttr($abstractNum,'abstractNumId');
                $numInfo = ['lvls'=>[],'abstractNumId'=>$abstractNumId,'text'=>[]];
                $lvls = $abstractNum->getElementsByTagName('lvl');
                foreach($lvls as $lvl) {
                    $ilvl = intval($this->getAttr($lvl, 'ilvl'));
                    $start = intval($this->getVal($lvl, 'start'));
                    $numFmt = $this->getVal($lvl, 'numFmt');
                    $lvlText = $this->getVal($lvl, 'lvlText');
                    $numInfo['lvls'][$ilvl] = ['lvl'=>$lvl,'start'=>$start,'isLgl'=>$this->getExist($lvl,'isLgl'),'numFmt'=>$numFmt,'lvlText'=>$lvlText];
                }

                $this->abstractNums[$abstractNumId] = $numInfo;
            }
        }

        return $this->abstractNums[$this->numIdToAbstractNumId[$numId][0]];
    }

    public function getText($numId,$ilvl) {
        static $numIdRecored = [];
        $num = $this->getAbstractNumById($numId);
        $lvlInfo = $num['lvls'][$ilvl];
        $lvl = $lvlInfo['lvl'];
        
        if(is_null($lvl)) {
            return '';
        }

        $abstractNumId = $num['abstractNumId'];
        $start = $lvlInfo['start'];
        $numFmt = $lvlInfo['numFmt'];
        $lvlText = $lvlInfo['lvlText'];

        if(!isset($numIdRecored[$numId])) {
            $numIdIsFirst = true;
        }else{
            $numIdIsFirst = false;
        }
        $numIdRecored[$numId] = true;
        if($numIdIsFirst === true) {
            $numNode = $this->numIdToAbstractNumId[$numId][1];
            $lvlOverrides = $numNode->getElementsByTagName('lvlOverride');
            if($lvlOverrides->length > 0) {
                foreach($lvlOverrides as $lvlOverride) {
                    $ilvlTemp = $this->getAttr($lvlOverride,'ilvl');
                    $startOverrideTemp = $this->getVal($lvlOverride,'startOverride');
                    $this->abstractNums[$abstractNumId]['text'][$ilvlTemp]['index'] = $startOverrideTemp;
                }
            }
        }

        foreach($this->abstractNums[$abstractNumId]['text'] as $ilvlTemp => $val) {
            if($ilvlTemp > $ilvl) {
                unset($this->abstractNums[$abstractNumId]['text'][$ilvlTemp]);
            }
        }

        if(!isset($this->abstractNums[$abstractNumId]['text'][$ilvl])) {
            $this->abstractNums[$abstractNumId]['text'][$ilvl] = [];
            $this->abstractNums[$abstractNumId]['text'][$ilvl]['index'] = $start;
        }

        $this->abstractNums[$abstractNumId]['text'][$ilvl]['text'] = $this->getTextByIndex($numFmt,$this->abstractNums[$abstractNumId]['text'][$ilvl]['index']);
        $text = preg_replace_callback('/\%(\d+)/i',function($match) use($abstractNumId,$num,$ilvl){
            $ilvlTemp = $match[1]-1;
            if(!isset($this->abstractNums[$abstractNumId]['text'][$ilvlTemp])) {
                $this->abstractNums[$abstractNumId]['text'][$ilvlTemp] = [];
                $this->abstractNums[$abstractNumId]['text'][$ilvlTemp]['index'] = $num['lvls'][$ilvlTemp]['start'];
                $this->abstractNums[$abstractNumId]['text'][$ilvlTemp]['text'] = $this->getTextByIndex($num['lvls'][$ilvlTemp]['numFmt'],$this->abstractNums[$abstractNumId]['text'][$ilvlTemp]['index']);
                $this->abstractNums[$abstractNumId]['text'][$ilvlTemp]['index']++;
            }

            if($num['lvls'][$ilvl]['isLgl'] === true) {
                return $this->getTextByIndex($num['lvls'][$ilvl]['numFmt'],$this->abstractNums[$abstractNumId]['text'][$ilvl]['index']);
            }else{
                return $this->abstractNums[$abstractNumId]['text'][$ilvlTemp]['text'];
            }
        },$lvlText);

        $this->abstractNums[$abstractNumId]['text'][$ilvl]['index']++;

        return $text;
    }

    private function getTextByIndex($numFmt,$index) {
        //$enum = array('bullet', 'decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter');
        switch($numFmt) {
            case 'decimal':
                return $index;
                break;
            case 'upperLetter':
                return chr(64+$index);
                break;
            case 'lowerLetter':
                return chr(96+$index);
                break;
            case 'upperRoman':
                $nf = new \NumberFormatter('@numbers=roman', \NumberFormatter::DECIMAL);
                return $nf->format($index);
            case 'lowerRoman':
                $nf = new \NumberFormatter('@numbers=roman', \NumberFormatter::DECIMAL);
                return strtolower($nf->format($index));
            case 'chineseCountingThousand':
                $nf = new \NumberFormatter('zh_CN', \NumberFormatter::SPELLOUT);
                return $nf->format($index);
                break;
        }
    }
}
