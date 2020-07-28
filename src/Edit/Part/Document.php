<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Document extends PartBase
{
    public $commentsblocks;
    public $charts = [];
    public $blocks = [];
    private $anchors = [];
    private $hyperlinkParentNodeArr = [];
    
    public function __construct($word,\DOMDocument $DOMDocument,$blocks = []) {
        parent::__construct($word);
        
        $this->DOMDocument = $DOMDocument;
        $this->commentsblocks = $blocks;
        
        $this->initNameSpaces();
        $this->initLevelToAnchor();
        $this->blocks = $this->initCommentRange();
    }
    
    private function initLevelToAnchor() {
        $sdtContent = $this->DOMDocument->getElementsByTagName('sdtContent')->item(0);
        if(is_null($sdtContent)) {
            $sdtContent = $this->DOMDocument;
        }
        $hyperlinks = $sdtContent->getElementsByTagName('hyperlink');
        $this->hyperlinkParentNodeArr = [];
        foreach($hyperlinks as $hyperlink) {
            $parentNode = $hyperlink->parentNode;
            $this->hyperlinkParentNodeArr[$this->getAttr($hyperlink, 'anchor')] = $parentNode;
        }
        
        
        $pPrs = $this->DOMDocument->getElementsByTagName('pPr');
        
        foreach ($pPrs as $pPr) {
            $pStyle = $pPr->getElementsByTagName('pStyle');
            if($pStyle->length > 0) {
                $pStyle = $pStyle->item(0);
                $val = intval($this->getAttr($pStyle, 'val'));
                if(isset($this->anchors[$val])) {
                    continue;
                }
                
                $bookmarkStarts = $pPr->parentNode->getElementsByTagName('bookmarkStart');
                foreach($bookmarkStarts as $bookmarkStart) {
                    if(isset($this->anchors[$val])) {
                        continue;
                    }
                    $name = $this->getAttr($bookmarkStart, 'name');
                    if(isset($this->hyperlinkParentNodeArr[$name])) {
                        $hyperlinkParentNode = $this->hyperlinkParentNodeArr[$name];
                        $copy = clone $hyperlinkParentNode;
                        $fldChars = $copy->getElementsByTagName('fldChar');
                        foreach($fldChars as $fldChar) {
                            $r = $fldChar->parentNode;
                            $p = $r->parentNode;
                            
                            if($p->localName === 'hyperlink') {
                                continue;
                            }
                            
                            $this->markDelete($r);
                        }
                        
                        $instrTexts = $copy->getElementsByTagName('instrText');
                        foreach($instrTexts as $instrText) {
                            $r = $instrText->parentNode;
                            $p = $r->parentNode;
                            
                            if($p->localName === 'hyperlink') {
                                continue;
                            }
                            
                            $this->markDelete($r);
                        }
                        
                        $this->anchors[$val] = $copy;
                    }
                }
            }
        }
    }
    
    /**
     * 
     * @param string $name
     * @param string|array $value 例如：
     * [
            [
            'text'=>1,
            'type'=>MDWORD_BREAK,
            ],
            [
            'text'=>'value blue',
            'style'=>'blue',
            'type'=>MDWORD_TEXT,
            ],
            [
            'text'=>'value red',
            'style'=>'red',
            'type'=>MDWORD_TEXT,
            ],
            'value plain'
        ];
     * @param string $type
     */
    public function setValue($name,$value,$type='text') {
        $blocks = $this->getBlocks($name);
        
        if(empty($blocks)) {
            $this->word->log->writeLog('not find name! name: '.$name);
        }
        
        foreach($blocks as $block) {
            $this->update($block,$name,$value,$type);
        }
    }
    
    public function updateToc() {
        $titles = $this->getTitles();
        if(empty($titles)) {
            return ;
        }
        $sdtContent = $this->DOMDocument->getElementsByTagName('sdtContent')->item(0);
        if(is_null($sdtContent)) {
            return ;
        }
        
        foreach($titles as $index => $title) {
            $anchor = $title['anchor'][0]['name'];
            $copy = $title['copy'];
            $copy->getElementsByTagName('t')->item(0)->nodeValue = $title['text'];
            
            if($tItme = $copy->getElementsByTagName('t')->item(1)) {
                $tItme->nodeValue = '';
            }
            $hyperlink = $copy->getElementsByTagName('hyperlink')->item(0);
            $this->setAttr($hyperlink, 'anchor', $title['anchor'][0]['name']);
            
            $instrTexts = $copy->getElementsByTagName('instrText');
            $instrTexts->item(0)->nodeValue = " PAGEREF $anchor \h ";
            
            if($index === 0) {//first add specail node
                $mixed = [
                    'r'=>[
                        'childs'=>[
                            'fldChar'=>[
                                'fldCharType'=>'begin',
                            ]
                        ],
                    ]
                ];
                $fldCharBegin = $this->creatNode($mixed);
                $this->insertBefore($fldCharBegin, $hyperlink);
                
                $mixed = [
                    'r'=>[
                        'childs'=>[
                            'instrText'=>[
                                'xml:space'=>'preserve',
                                'text'=>' TOC \o "1-3" \h \z \u ',
                            ]
                        ],
                    ]
                ];
                $fldCharPreserve = $this->creatNode($mixed);
                $this->insertBefore($fldCharPreserve, $hyperlink);
                
                $mixed = [
                    'r'=>[
                        'childs'=>[
                            'fldChar'=>[
                                'fldCharType'=>'separate',
                            ]
                        ],
                    ]
                ];
                $fldCharSeparate = $this->creatNode($mixed);
                $this->insertBefore($fldCharSeparate, $hyperlink);
            }
            
            
            $sdtContent->insertBefore($copy,$sdtContent->lastChild);
        }
        
        $fldChars = $sdtContent->getElementsByTagName('fldChar');
        $fldCharsCount = [];
        foreach($fldChars as $fldChar) {
            $fldCharType = $this->getAttr($fldChar, 'fldCharType');
            if(empty($this->getAttr($fldChar->parentNode, 'md',null))) {
                $fldCharsCount[$fldCharType]++;
            }
        }
        
        if($fldCharsCount['begin'] > $fldCharsCount['end']) {
            $mixed = [
                'r'=>[
                    'childs'=>[
                        'fldChar'=>[
                            'fldCharType'=>'end',
                        ]
                    ],
                ]
            ];
            $fldCharEnd = $this->creatNode($mixed);
            $copy->appendChild($fldCharEnd);
        }
        
        
        foreach($this->hyperlinkParentNodeArr as $hyperlinkParentNode) {
            $this->markDelete($hyperlinkParentNode);
        }
    }
    
    private function getTitles() {
        $pPrs = $this->DOMDocument->getElementsByTagName('pPr');
        
        $titles = [];
        foreach ($pPrs as $pPr) {
            $pStyle = $pPr->getElementsByTagName('pStyle');
            if($pStyle->length > 0) {
                $pStyle = $pStyle->item(0);
                $val = intval($this->getAttr($pStyle, 'val'));
                if(isset($this->anchors[$val])) {
                    $anchorInfo = $this->updateBookMark($pPr->parentNode);
                    $anchorNode = $this->anchors[$val];
                    $copy = clone $anchorNode;
                    
                    $titles[] = ['copy'=>$copy,'anchor'=>$anchorInfo,'text'=>$pPr->parentNode->textContent];
                }
            }
        }
        
        return $titles;
    }
    
    private function getTocLevels() {
        $instrText = $this->DOMDocument->getElementsByTagName('instrText');
        if(!$instrText->length) {
           return []; 
        }
        
        $instrText = $instrText->item(0);
        $nodeValue = $instrText->nodeValue;
        preg_match('/TOC[\s\S]+?"(\d+)\-(\d+)\"/i', $nodeValue,$match);
        
        $begin = intval($match[1]);
        $end = intval($match[2]);
        
        $levels = [];
        if($begin > 0 && $end > 0 && $end > $begin) {
            for($begin;$begin<=$end;$begin++) {
                $levels[] = $begin;
            }
        }
        
        return $levels;
    }
    
    private function getStyle($name='',$type=MDWORD_TEXT) {
        static $styles = [];
        
        $stylekey = $name.$type;
        
        if(isset($styles[$stylekey])) {
            return $styles[$stylekey];
        }
        
        $nodeIdxs = $this->getBlocks($name);
        if(!isset($nodeIdxs[0])) {
            return null;
        }
        
        $nodeIdxs = $nodeIdxs[0];
        
        switch ($type) {
            case MDWORD_TEXT:
                $targetNode = $this->getTarget($nodeIdxs,'r');
                if($rPrs = $targetNode->getElementsByTagName('rPr')) {
                    $rPr = $rPrs->item(0);
                }
                
                $styles[$stylekey] = $rPr;
                return $rPr;
                break;
            case MDWORD_IMG:
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,$nextNodeCount,'drawing');
                $styles[$stylekey] = $targetNode;
                return $targetNode;
        }
    }
    
    public function update($nodeIdxs,$name,$value,$type) {
        static $count = 0;
        $count++;
        if($name === 'content') {
            var_dump($nodeIdxs);exit;
        }
        switch ($type) {
            case 'text':
                $targetNode = $this->getTarget($nodeIdxs,'r',function($node) {
                    $t = $node->getElementsByTagName('t');
                    if($t->length > 0) {
                        return true;
                    }else{
                        return false;
                    }
                });
                if(is_null($targetNode)) {
                    break;
                }
                
                if(is_array($value)) {
                    foreach($value as $valueArr) {
                        $copy = clone $targetNode;
                        if(is_array($valueArr)) {
                            switch ($valueArr['type']){
                                case MDWORD_BREAK:
                                    $valueArr['text'] = intval($valueArr['text']);
                                    $copyP = $this->updateMDWORD_BREAK($targetNode->parentNode,$valueArr['text'],false);
                                    $this->markDelete($targetNode);
                                    
                                    //get first r include t
                                    $rs  = $copyP->getElementsByTagName('r');
                                    foreach($rs as $r) {
                                        $t = $r->getElementsByTagName('t');
                                        if($t->length > 0) {
                                            $targetNode = $r;
                                            break;
                                        }
                                    }
                                    
                                    $this->removeMarkDelete($targetNode);
                                    break;
                                case MDWORD_PAGE_BREAK:
                                    break;
                                case MDWORD_LINK:
//                                     echo $this->DOMDocument->saveXML($copy);exit;
//                                     var_dump($copy);exit;
                                    if(!is_null($valueArr['text'])) {
                                        $copy->getElementsByTagName('t')->item(0)->nodeValue= $valueArr['text'];
                                    }
                                    $this->insertBefore($copy, $targetNode);
                                    $this->markDelete($targetNode);
                                    $this->updateMDWORD_LINK($copy, $copy, $valueArr['link']);
                                    break;
                                case MDWORD_IMG:
                                    $drawing = $this->getStyle($valueArr['style'],MDWORD_IMG);
                                    $copyDrawing = clone $drawing;
                                    
                                    $refInfo = $this->updateRef($valueArr['text'],null,MDWORD_IMG);
                                    $rId = $refInfo['rId'];
                                    $imageInfo = $refInfo['imageInfo'];
                                    
                                    $blip = $copyDrawing->getElementsByTagName('blip')->item(0);
                                    $this->setAttr($blip, 'embed', $rId,'r');
                                    
                                    $orgCx = intval($imageInfo[0]*9530);
                                    $orgCy = intval($imageInfo[1]*9530);
                                    
                                    $extents = $copyDrawing->getElementsByTagName('extent');
                                    foreach($extents as $extent) {
                                        $this->setAttr($extent, 'cx', $orgCx, null);
                                        $this->setAttr($extent, 'cy', $orgCy, null);
                                    }
                                    
                                    if($spPr = $copyDrawing->getElementsByTagName('spPr')->item(0)) {
                                        $exts = $spPr->getElementsByTagName('ext');
                                        foreach($exts as $extent) {
                                            $this->setAttr($extent, 'cx', $orgCx, null);
                                            $this->setAttr($extent, 'cy', $orgCy, null);
                                        }
                                    }
                                    
                                    $targetNode->getElementsByTagName('t')->item(0)->nodeValue= '';
                                    $targetNode->appendChild($copyDrawing);
                                    
                                    $copyP = $this->updateMDWORD_BREAK($targetNode->parentNode,1,false,['drawing']);
                                    $targetNode = $copyP->getElementsByTagName('r')->item(0);
                                    $this->removeMarkDelete($targetNode);
                                    break;
                                default:
                                    $rPr = $this->getStyle($valueArr['style']);
                                    if(!is_null($rPr)) {
                                        $rPrCopy = clone $rPr;
                                        $rPrOrg = $copy->getElementsByTagName('rPr')->item(0);
                                        if(is_null($rPrOrg)) {// rPr insert before t
                                            $rPrOrg = $copy->getElementsByTagName('t')->item(0);
                                            $this->insertBefore($rPrCopy, $rPrOrg);
                                        }else{
                                            $this->insertBefore($rPrCopy, $rPrOrg);
                                            $this->markDelete($rPrOrg);
                                        }
                                    }
                                    $copy->getElementsByTagName('t')->item(0)->nodeValue= $valueArr['text'];
                                    $this->insertBefore($copy, $targetNode);
                                    break;
                            }
                        }else{
                            $copy->getElementsByTagName('t')->item(0)->nodeValue= $valueArr;
                            $this->insertBefore($copy, $targetNode);
                        }
                        
                    }
                    
                }else{
                    $copy = clone $targetNode;
                    $copy->getElementsByTagName('t')->item(0)->nodeValue= $value;
                    $this->insertBefore($copy, $targetNode);
                }
                
                $this->markDelete($targetNode);
                break;
            case 'image':
                $rids = $this->getRidByMd5($name);
                if(empty($rids)) {
                    $this->word->log->writeLog('not find image by md5! md5: '.$name);
                }
                foreach($rids as $rid) {
                    $this->updateRef($value,$rid);
                }
                break;
            case 'excel':
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,$nextNodeCount,'drawing');
                if(!is_null($targetNode)) {
                    $rid = $this->getAttr($targetNode->getElementsByTagName('chart')->item(0), 'id', 'r');
                    $this->initChart($rid);
                    $value = $this->charts[$rid]->excel->preDealDatas($value);
                    $this->charts[$rid]->excel->changeExcelValues($value);
                    $this->charts[$rid]->chartRelUpdateByType($value,'str');
                    $this->charts[$rid]->chartRelUpdateByType($value,'num');
                }
                break;
            case 'cloneP':
                $p = $lastNode = $this->getParentToNode($beginNode,'p');
                $needCloneNodes = [$p];
                for($i=1;$i<=$value;$i++) {
                    foreach($needCloneNodes as $targetNode) {
                        $copy = clone $targetNode;
                        $this->updateCommentsId($copy, $i, $value);
                        if($nextSibling = $lastNode->nextSibling) {
                            $this->insertBefore($copy, $nextSibling);
                        }else{
                            $parentNode = $lastNode->parentNode;
                            $parentNode->appendChild($copy);
                        }
                        
                        $lastNode = $copy;
                    }
                }
                
                foreach($needCloneNodes as $targetNode) {
                    $this->updateCommentsId($targetNode, 0);
                }
                break;
            case 'cloneTo':
                switch($value['type']) {
                    case MDWORD_DELETE:
                        foreach($nodeIdxs as $nodeIdx) {
                            $p = $this->domList[$nodeIdx];
                            if(!is_null($p)) {
                                $this->markDelete($p);
                            }
                        }
                        break;
                }
                break;
            case 'clone':
                if($this->isTc($this->domList[$nodeIdxs[0]])) {
                    $nodeIdxs = [$this->domList[$nodeIdxs[0]]->parentNode->idxBegin];
                }
                
                $lastNodeIdx = end($nodeIdxs);
                for($i=1;$i<$value;$i++) {
                    foreach($nodeIdxs as $nodeIdx) {
                        $lastNodeIdx = $this->cloneNode($nodeIdx,$lastNodeIdx,$name,$i);
                    }
                }
                
                //刷新被克隆对象
                foreach($nodeIdxs as $nodeIdx) {
                    $lastNodeIdx = $this->cloneNode($nodeIdx,$lastNodeIdx,$name,0);
                }
                break;
            case 'delete':
                if($value == 'p') {
                    foreach($nodeIdxs as $nodeIdx) {
                        $p = $this->getParentToNode($nodeIdx,'p');
                        if(!is_null($p)) {
                            $this->markDelete($p);
                        }
                    }
                }elseif($value == 'tr') {//to-do
                    foreach($nodeIdxs as $nodeIdx) {
                        $tr = $this->getParentToNode($nodeIdx,'tr');
                        if(!is_null($tr)) {
                            $this->markDelete($tr);
                        }
                    }
                }
                break;
            case 'break':
                $p = $lastNode = $this->getParentToNode($nodeIdxs[0],'p');
//                 var_dump($p);exit;
                $this->updateMDWORD_BREAK($p,$value,true);
                break;
            case 'breakpage':
                $p = $lastNode = $this->getParentToNode($beginNode,'p');
                $childNodes = $p->childNodes;
                foreach($childNodes as $childNode) {
                    $this->markDelete($childNode);
                }
                
                $needCloneNodes = [$p];
                
                $mixed = [
                    'r'=>[
                        'childs'=>[
                            'br'=>[
                                'type'=>'page',
//                                 'xmlns:default'=>null,
                            ]
                        ],
                    ]
                ];
                $breakpage = $this->creatNode($mixed);
                for($i=1;$i<=$value;$i++) {
                    foreach($needCloneNodes as $targetNode) {
                        $copy = clone $targetNode;
//                         var_dump($copy,$breakpage);exit;
                        $copy->appendChild($breakpage);
                        $this->updateCommentsId($copy, $i, $value);
                        if($nextSibling = $lastNode->nextSibling) {
                            $this->insertBefore($copy,$nextSibling);
                        }else{
                            $parentNode = $lastNode->parentNode;
                            $parentNode->appendChild($copy);
                        }
                        
                        $lastNode = $copy;
                    }
                }
                break;
            case 'link':
                $this->updateMDWORD_LINK($beginNode, $endNode, $value);
                break;
            default:
                break;
        }
    }
    
    private function cloneNode($nodeIdx,$endNodeIdx,$name,$idx) {
        $node = $this->domList[$nodeIdx];
        if($idx === 0) {
            $begin = $node->idxBegin;
            $end = $node->idxEnd;
            $offset = 0;
            for($i = $begin; $i <= $end; $i++) {
                if(isset($this->domIdxToName[$i])) {
                    $cloneNodeIdx = $i+$offset;
                    $nameTemps = $this->domIdxToName[$i];
                    foreach($nameTemps as $key => $nameTemp) {
                        $newName = $nameTemp[1].'#'.$idx;
                        $nameTemps[$key] = [$nameTemp[0],$newName];
                        $this->blocks[$newName][$nameTemp[0]][] = $cloneNodeIdx;
                    }
                    $this->domIdxToName[$cloneNodeIdx] = $nameTemps;
                }
            }
            
            return $node->idxBegin;
        }else{
            $cloneNode = clone $node;
            $begin = $node->idxBegin;
            $end = $node->idxEnd;
            $baseIndex = $this->treeToList(null);
            $this->treeToList($cloneNode);
            $offset = $baseIndex - $begin;
            
            for($i = $begin; $i <= $end; $i++) {
                if(isset($this->domIdxToName[$i])) {
                    $cloneNodeIdx = $i+$offset;
                    if(isset($this->idxExtendIdxs[$i])) {
                        $this->idxExtendIdxs[$cloneNodeIdx] = $this->idxExtendIdxs[$i];
                    }
                    
                    $nameTemps = $this->domIdxToName[$i];
                    foreach($nameTemps as $key => $nameTemp) {
                        $newName = $nameTemp[1].'#'.$idx;
                        $nameTemps[$key] = [$nameTemp[0],$newName];
                        $this->blocks[$newName][$nameTemp[0]][] = $cloneNodeIdx;
                    }
                    $this->domIdxToName[$cloneNodeIdx] = $nameTemps;
                }
            }
            
            
            $this->insertAfter($cloneNode, $this->domList[$endNodeIdx]);
            return $baseIndex;
        }
    }
    
    private function isTc($item) {
        if($item->localName === 'tc') {
            return true;
        }else{
            return false;
        }
    }
    
    private function updateBookMark($item) {
        static $maxId = 10000;
        //start
        if($item->localName === 'bookmarkStart') {
            $bookmarkStarts = [$item];
        }else{
            $bookmarkStarts = $item->getElementsByTagName('bookmarkStart');
        }
        
        $ids = [];
        $infos = [];
        foreach($bookmarkStarts as $key => $bookmarkStart) {
            $maxId = $maxId+1;
            $name = '_Toc'.$maxId;
            $ids[$key] = $maxId;
            $this->setAttr($bookmarkStart, 'id', $maxId);
            $this->setAttr($bookmarkStart, 'name', $name);
            $infos[] = ['id'=>$maxId,'name'=>$name];
        }
        
        //end
        if($item->localName === 'bookmarkEnd') {
            $bookmarkEnds = [$item];
        }else{
            $bookmarkEnds = $item->getElementsByTagName('bookmarkEnd');
        }
        foreach($bookmarkEnds as $key => $bookmarkEnd) {
            $this->setAttr($bookmarkEnd, 'id', $ids[$key]);
        }
        
        
        if(empty($infos)) {//no include book market,insert one
            $rs = $item->getElementsByTagName('r');
            if($rs->length > 0) {
                $maxId = $maxId+1;
                $name = '_Toc'.$maxId;
                $mixed = [
                    'bookmarkStart'=>[
                        'id'=>$maxId,
                        'name'=>$name,
                    ],
                ];
                $bookmarkStart = $this->creatNode($mixed);
                $this->insertBefore($bookmarkStart, $rs->item(0));
                
                $mixed = [
                    'bookmarkEnd'=>[
                        'id'=>$maxId
                    ],
                ];
                $bookmarkEnd = $this->creatNode($mixed);
                $this->insertAfter($bookmarkStart, $rs->item($rs->length-1));
                
                $infos[] = ['id'=>$maxId,'name'=>$name];
            }
        }
        
        return $infos;
    }
    
    
    private function updateCommentsId($item,$id,$value='') {
        //start
        if($item->localName === 'commentRangeStart') {
            $commentRangeStarts = [$item];
        }else{
            $commentRangeStarts = $item->getElementsByTagName('commentRangeStart');
        }
        foreach($commentRangeStarts as $commentRangeStart) {
            $oldId = $this->getAttr($commentRangeStart, 'id');
            if($name = $this->commentsblocks[$oldId]) {
                $name .= '#'.$id;
            }else{
                $name = 'TEMP#'.$id;
            }
            $this->commentsblocks[$name] = $name;
            $this->blocks[$name] = $this->blocks[$this->commentsblocks[$oldId]];
            $this->blocks[$name][0][0] = $commentRangeStart;
            $this->setAttr($commentRangeStart, 'id', $name);
        }
        
        //end
        if($item->localName === 'commentRangeEnd') {
            $commentRangeEnds = [$item];
        }else{
            $commentRangeEnds = $item->getElementsByTagName('commentRangeEnd');
        }
        foreach($commentRangeEnds as $commentRangeEnd) {
            $oldId = $this->getAttr($commentRangeEnd, 'id');
            if($name = $this->commentsblocks[$oldId]) {
                $name .= '#'.$id;
            }else{
                $name = 'TEMP#'.$id;
            }
            $this->blocks[$name][0][1] = $commentRangeEnd;
            $this->setAttr($commentRangeEnd, 'id', $name);
        }
    }
    
    private function getTarget($nodeIdxs,$type='r' ,$checkCallBack = null) {
        $find = false;
        foreach($nodeIdxs as $nodeIdx) {
            $node = $this->domList[$nodeIdx];
            $this->markDelete($node);
            if($find) {
                continue;
            }
            
            if($node->localName === $type && (is_null($checkCallBack) || $checkCallBack($node))) {
                $targetNode = $node;
                $find = true;
            }else{//todo
                $rs = $node->getElementsByTagName($type);
                foreach($rs as $r) {
                    if(is_null($targetNode) && (is_null($checkCallBack) || $checkCallBack($r))) {
                        $targetNode = $r;
                        $this->removeMarkDelete($node);
                    }else{//delete sub pre node
                        foreach($targetNode->parentNode->childNodes as $item) {
                            $this->markDelete($item);
                        }
                    }
                }
            }
        }
        
        $this->removeMarkDelete($targetNode);
        
        return $targetNode;
    }
    
    private function getParentToNode($beginNodeIndex,$type='p') {
        $parentNode = $this->domList[$beginNodeIndex];
        while($parentNode->localName != $type && !is_null($parentNode)) {
            $parentNode = $parentNode->parentNode;
        }
        
        return $parentNode;
    }
    
    private function initChart($rid='') {
        if(is_null($this->rels)) {
            $this->initRels();
        }
        
        if(!isset($this->charts[$rid])) {
            $target = $this->rels->getTarget($rid);
            
            $dom = null;
            foreach($this->word->parts[13] as $chart) {
                if($chart['PartName'] === $target) {
                    $dom = $chart['DOMElement'];
                }
                
            }
            
            if(is_null($dom)) {
                $this->charts[$rid] = new Charts($this->word, $this->word->getXmlDom($target));
            }else{
                $this->charts[$rid] = new Charts($this->word, $dom);
            }
            $this->charts[$rid]->partName = $target;
            $this->charts[$rid]->initRels(13);
            
            $this->charts[$rid]->excel = new Excel($this->word,$this->charts[$rid]->rels->getTarget());
        }
    }
    
    private function updateRef($file,$rid=null,$type=MDWORD_IMG) {
        if(is_null($this->rels)) {
            $this->initRels();
        }
        
        if(is_null($rid)) {
            return $this->rels->insert($file,$type);
        }else{
            return $this->rels->replace($rid,$file);
        }
    }
    
    private function getRidByMd5($md5) {
        if(is_null($this->rels)) {
            $this->initRels();
        }
        
        return $this->rels->imageMd5ToRid[$md5];
    }
    
    private function deleteBlock($block) {
        $beginNode = $block[0]['node'];
        $endNode = $block[1]['node'];
        
        $parentNode = $beginNode->parentNode;
        
        if(!is_null($beginNode)) {
            $parentNode->removeChild($beginNode);
        }
        if(!is_null($endNode)) {
            $parentNode->removeChild($endNode);
        }
    }
    
    private function getBlocks($name) {
        if(!isset($this->blocks[$name])) {
            return [];
        }
        
        $blocks = [];
        foreach($this->blocks[$name] as $key => $indexs) {
            $blocks[$key] = [];
            foreach($indexs as $index) {
                $blocks[$key][] = $index;
                if(isset($this->idxExtendIdxs[$index])) {
                    foreach($this->idxExtendIdxs[$index] as $index2) {
                        $blocks[$key][] = $index2;
                    }
                }
            }
        }
        
        return $blocks;
    }
    
    
    private function updateMDWORD_BREAK($p,$count=1,$replace=true,$needDelTags = []) {
        if($replace === true) {
            $count -= 1;
            $copyP = $p;
        }else{
            $copyP = clone $p;
        }
        
        $childNodes = $copyP->childNodes;
        foreach($childNodes as $childNode) {
            if($childNode->localName === 'pPr') {
                continue;
            }
            
            $this->markDelete($childNode);
        }
        
        foreach($needDelTags as $needDelTag) {
            $nodes = $copyP->getElementsByTagName($needDelTag);
            foreach($nodes as $node) {
                $node->parentNode->removeChild($node);
            }
        }
        
        $indexs = [];
        for($i=0;$i<$count;$i++) {
            $copy = clone $copyP;
            $baseIndex = $this->treeToList(null);
            $this->treeToList($copy);
            $indexs[] = $baseIndex;
            $this->insertAfter($copy, $p);
            $p = $copy;
        }
        
        if(count($indexs) > 0) {
            $this->idxExtendIdxs[$copyP->idxBegin] = $indexs;
        }
        
        return $p;
    }
    
    private function updateMDWORD_LINK($beginNode,$endNode,$link) {
        $mixed = [
            'r'=>[
                'childs'=>[
                    'fldChar'=>[
                        'fldCharType'=>'begin',
                    ]
                ],
            ],
        ];
        $hyperlinkNodeBegin = $this->creatNode($mixed);
        
        $mixed = [
            'r'=>[
                'childs'=>[
                    'instrText'=>[
                        'xml:space'=>'preserve',
                        'text'=>' HYPERLINK "'.$link.'" '
                    ]
                ],
            ],
        ];
        $hyperlinkNodePreserve = $this->creatNode($mixed);
        
        $mixed = [
            'r'=>[
                'childs'=>[
                    'fldChar'=>[
                        'fldCharType'=>'separate',
                    ]
                ],
            ],
        ];
        $hyperlinkNodeSeparate = $this->creatNode($mixed);
        
        $mixed = [
            'r'=>[
                'childs'=>[
                    'fldChar'=>[
                        'fldCharType'=>'end'
                    ]
                ],
            ],
        ];
        $hyperlinkNodeEnd = $this->creatNode($mixed);
        
        $this->insertBefore($hyperlinkNodeBegin, $beginNode);
        $this->insertBefore($hyperlinkNodePreserve, $beginNode);
        $this->insertBefore($hyperlinkNodeSeparate, $beginNode);
        $this->insertAfter($hyperlinkNodeEnd, $endNode);
    }

    public function setChartRel($relArr){
        $this->initChartRels($relArr);
    }

    public function getDocumentChart(){
        $chartsP = $this->DOMDocument->getElementsByTagName('p');
        foreach($chartsP as $chart){
            $length = $chart->getElementsByTagName('chart')->length;
            if($length>0){
                $documentChart[] = $chart;
            }
        }
        return $documentChart;
    }

    public function setDocumentChart($name,$documentChart){
        if(is_null($this->rels)) {
            $this->initRels();
        }
        $chartRidArr = $this->rels->setNewChartRels(count($documentChart));
        $blocks = $this->getBlocks($name);
        foreach($blocks as $block){
            $beginNode = $block[0];
            $endNode = $block[1];
       
            $traces = $block[2];
            $parentNodeCount = $traces['parentNodeCount'];
            $nextNodeCount = $traces['nextNodeCount'];

            $targetNode = $this->getParentToNode($beginNode,'p');

            foreach($documentChart as $key=>$chart){
                $this->setAttr($chart->getElementsByTagName('chart')->item(0), 'id', 'rId'.$chartRidArr[$key],'r');
                if($res = $this->DOMDocument->importNode($chart,true)){
                    $this->insertBefore($res,$targetNode);
                 }
            }
        }
        if(!is_null($targetNode->parentNode)){
            $targetNode->parentNode->removeChild($targetNode); 
        }
    }
}
