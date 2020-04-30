<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Document extends PartBase
{
    public $commentsblocks;
    public $charts = [];
    public $blocks = [];
    
    public function __construct($word,\DOMDocument $DOMDocument,$blocks = []) {
        parent::__construct($word);
        
        $this->DOMDocument = $DOMDocument;
        $this->commentsblocks = $blocks;
        
        $this->initNameSpaces();
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
        foreach($blocks as $block) {
            $this->update($block,$value,$type);
        }
    }
    
    public function updateToc() {
        //get title levels
        $levels = $this->getTocLevels();
        
        $titles = $this->getTitles($levels);
        
        if(empty($titles)) {
            return ;
        }
        $sdtContent = $this->DOMDocument->getElementsByTagName('sdtContent')->item(0);
        $hyperlinks = $sdtContent->getElementsByTagName('hyperlink');
        
        $hyperlinkParentNodeArr = [];
        foreach($hyperlinks as $key => $hyperlink) {
            $parentNode = $hyperlink->parentNode;
            $hyperlinkParentNodeArr[$this->getAttr($hyperlink, 'anchor')] = $parentNode;
        }
        
        $index = 0;
        foreach($titles as $anchor => $title) {
            $anchorOrg = $title['orgname'];
            
            if(isset($hyperlinkParentNodeArr[$anchorOrg])) {
                $index++;
                $p = $hyperlinkParentNodeArr[$anchorOrg];
                $copy = clone $p;
                $copy->getElementsByTagName('t')->item(0)->nodeValue = $title['text'];
                
                if($tItme = $copy->getElementsByTagName('t')->item(1)) {
                    $tItme->nodeValue = '';
                }
//                 var_dump($copy->getElementsByTagName('t')->item(1));exit;
                $hyperlink = $copy->getElementsByTagName('hyperlink')->item(0);
                $this->setAttr($hyperlink, 'anchor', $anchor);
                
                $instrTexts = $copy->getElementsByTagName('instrText');
                if($instrTexts.length > 1) {
                    if($index !== 1) {
                        $this->markDelete($instrTexts->item(0)->parentNode);
                    }
                    $instrTexts->item(1)->nodeValue = str_replace($anchorOrg, $anchor, $instrTexts->item(1)->nodeValue);
                }else{
                    $instrTexts->item(0)->nodeValue = str_replace($anchorOrg, $anchor, $instrTexts->item(0)->nodeValue);
                }
                
                $sdtContent->insertBefore($copy,$sdtContent->lastChild);
            }
        }
        
        if(!empty($copy)) {
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
        
        
        foreach($hyperlinkParentNodeArr as $hyperlinkParentNode) {
            $this->markDelete($hyperlinkParentNode);
        }
    }
    
    private function getTitles($levels=[]) {
        $pPrs = $this->DOMDocument->getElementsByTagName('pPr');
        
        $titles = [];
        foreach ($pPrs as $pPr) {
            $pStyle = $pPr->getElementsByTagName('pStyle');
            if($pStyle->length > 0) {
                $pStyle = $pStyle->item(0);
                $val = intval($this->getAttr($pStyle, 'val'));
                if(in_array($val, $levels)) {
                    $bookmarkStarts = $pPr->parentNode->getElementsByTagName('bookmarkStart');
                    foreach($bookmarkStarts as $bookmarkStart) {
                        $name = $this->getAttr($bookmarkStart, 'name');
                        $id = $this->getAttr($bookmarkStart, 'id');
                        $orgname = $this->getAttr($bookmarkStart, 'orgname');
                        if($orgname == '') {
                            $orgname = $name;
                        }
                        $titles[$name] = ['id'=>$id,'orgname'=>$orgname,'text'=>$pPr->parentNode->textContent];
                    }
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
        
        $blocks = $this->getBlocks($name);
        if(!isset($blocks[0])) {
            return null;
        }
        
        $block = $blocks[0];
        $beginNode = $block[0];
        $endNode = $block[1];
        
        $traces = $block[2];
        $parentNodeCount = $traces['parentNodeCount'];
        $nextNodeCount = $traces['nextNodeCount'];
        
        
        switch ($type) {
            case MDWORD_TEXT:
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,$nextNodeCount,'r');
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
    
    private function update($block,$value,$type) {
        $beginNode = $block[0];
        $endNode = $block[1];
        
        $traces = $block[2];
        $parentNodeCount = $traces['parentNodeCount'];
        $nextNodeCount = $traces['nextNodeCount'];
        switch ($type) {
            case 'text':
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,$nextNodeCount,'r',function($node) {
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
                                    $targetNode = $copyP->getElementsByTagName('r')->item(0);
                                    $this->removeMarkDelete($targetNode);
                                    break;
                                case MDWORD_PAGE_BREAK:
                                    break;
                                case MDWORD_LINK:
                                    $copy->getElementsByTagName('t')->item(0)->nodeValue= $valueArr['text'];
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
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,$nextNodeCount,'pict');
                if(is_null($targetNode)) {
                    break;
                }
                $rid = $this->getAttr($targetNode->getElementsByTagName('imagedata')->item(0), 'id', 'r');
                $this->updateRef($rid,$value);
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
            case 'clone':
                $parentNode = $beginNode;
                $targetNode = null;
                $needCloneNodes = [];
                for($i=0;$i<$parentNodeCount;$i++) {
                    $parentNode = $parentNode->parentNode;
                }
                
                
                if($this->isTc($parentNode)) {
                    $parentNode = $parentNode->parentNode;
                    $nextNodeCount = 0;
                }
                $needCloneNodes[] = $lastNode = $parentNode;
                
                $nextSibling = $parentNode;
                for($i=0;$i<$nextNodeCount;$i++) {
                    $nextSibling = $nextSibling->nextSibling;
                    $needCloneNodes[] = $lastNode = $nextSibling;
                }
                
                for($i=1;$i<=$value;$i++) {
                    foreach($needCloneNodes as $targetNode) {
                        $copy = clone $targetNode;
                        $this->updateCommentsId($copy, $i, $value);
                        if($nextSibling = $lastNode->nextSibling) {
                            $parentNode = $nextSibling->parentNode;
                            $parentNode->insertBefore($copy,$nextSibling);
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
                
//                 var_dump($this->blocks);exit;
                break;
            case 'delete':
                if($value == 'p') {
                    $p = $this->getParentToNode($beginNode,'p');
                    $this->markDelete($p);
                }else{//to-do
                    
                }
                break;
            case 'break':
                $p = $lastNode = $this->getParentToNode($beginNode,'p');
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
    
    private function isTc($item) {
        if($item->localName === 'tc') {
            return true;
        }else{
            return false;
        }
    }
    
    private function updateBookMark($item,$id,$value='') {
        static $maxId = 10000;
        //start
        if($item->localName === 'bookmarkStart') {
            $bookmarkStarts = [$item];
        }else{
            $bookmarkStarts = $item->getElementsByTagName('bookmarkStart');
        }
        
        $ids = [];
        foreach($bookmarkStarts as $key => $bookmarkStart) {
            $maxId = $maxId+1;
            $name = '_Toc'.$maxId;
            $ids[$key] = $maxId;
            $this->setAttr($bookmarkStart, 'id', $maxId);
            if(!$this->hasAttr($bookmarkStart, 'orgname')) {
                $this->setAttr($bookmarkStart, 'orgname', $this->getAttr($bookmarkStart, 'name'));
            }
            $this->setAttr($bookmarkStart, 'name', $name);
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
    }
    
    
    private function updateCommentsId($item,$id,$value='') {
        //org clone not change bookmark's id and name
        if($id != 0) {
            $this->updateBookMark($item, $id, $value);
        }
        
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
    
    private function getTarget($beginNode,$endNode,$parentNodeCount,$nextNodeCount,$type='r',$checkCallBack = null) {
        $keepTags = ['bookmarkEnd'=>1,'bookmarkStart'=>1,'rPr'=>1];
        $parentNode = $beginNode;
        $targetNode = null;
        for($i=0;$i<=$parentNodeCount;$i++) {
            $nextSibling = $parentNode;
            
            if($i === $parentNodeCount) {//top parent
                $maxNext = $nextNodeCount;
            }else{
                $parentNode = $parentNode->parentNode;
                $maxNext = 0;
            }
            
            $j = 0;
            while($nextSibling = $nextSibling->nextSibling) {
//                 $this->markDelete($nextSibling);
                if($maxNext > 0 && ++$j > $maxNext) {
                    break 2;
                }
//                 if($nextSibling === $endNode) {
//                     break 2;
//                 }
                if(is_null($targetNode)) {
                    if($nextSibling->localName == $type && (is_null($checkCallBack) || $checkCallBack($nextSibling))) {//is target
                        $targetNode = $nextSibling;
                    }else{//sub find target
                        $rs = $nextSibling->getElementsByTagName($type);
                        foreach($rs as $r) {
                            if(is_null($targetNode) && (is_null($checkCallBack) || $checkCallBack($r))) {
                                $targetNode = $r;
                            }
                        }
                        
                        if(is_null($targetNode)){
                            if(!isset($keepTags[$nextSibling->localName])) {
                                $this->markDelete($nextSibling);
                            }
                        }else{//delete sub pre node
                            foreach($targetNode->parentNode->childNodes as $item) {
                                if($endNode === $item) {
                                    break 2;
                                }
                                
                                if($targetNode !== $item && !isset($keepTags[$item->localName])) {
//                                     var_dump($item);
                                    $this->markDelete($item);
                                }
                            }
//                             var_dump($targetNode->parentNode,$targetNode->parentNode->childNodes);exit;
                        }
                        
                        
                    }
                }
                else{
                    if(!isset($keepTags[$nextSibling->localName])) {
                        $this->markDelete($nextSibling);
                    }
                }
            }
            
        }
        
        if(is_null($targetNode)) {
            $targetNodesTemp = $beginNode->parentNode->getElementsByTagName($type);
            if($targetNodesTemp->length > 0) {
                $targetNode = $targetNodesTemp->item(0);
            }
        }
//         $targetNode
        $this->removeMarkDelete($targetNode);
        return $targetNode;
    }
    
    private function getParentToNode($beginNode,$type='p') {
        $parentNode = $beginNode;
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
        if(isset($this->blocks[$name])) {
            return $this->blocks[$name];
        }
        
        $this->blocks[$name] = [];
        $commentRangeStartItems = $this->DOMDocument->getElementsByTagName('commentRangeStart');
        foreach($commentRangeStartItems as $commentRangeStartItem) {
            $id = $this->getAttr($commentRangeStartItem, 'id');
            $nameTemp = $this->commentsblocks[$id];
            $commentRangeEndItem = $this->getCommentRangeEnd($this->DOMDocument,$id);
            $trace = $this->getRangeTrace($id,$commentRangeStartItem, $commentRangeEndItem);
            $this->blocks[$nameTemp][] = [$commentRangeStartItem,$commentRangeEndItem,$trace];
        }
        
        return $this->blocks[$name];
    }
    
    private function getRangeTrace($id,$commentRangeStartItem,$commentRangeEndItem) {
        $startParentNode = $parentNode = $commentRangeStartItem;
        $parentNodeCount = 0;
        while($parentNode = $parentNode->parentNode) {
            $commentRangeEndItem = $this->getCommentRangeEnd($parentNode,$id);
            if(is_null($commentRangeEndItem)) {
                $startParentNode = $parentNode;
            }else{
                break;
            }
            $parentNodeCount++;
        }
        
        $nextNodeCount = 0;
        if($parentNodeCount === 0) {
            $nextSibling = $commentRangeStartItem->nextSibling;
            $nextNodeCount++;
            while($nextSibling !== null && $nextSibling !== $commentRangeEndItem) {
                $nextSibling = $nextSibling->nextSibling;
                $nextNodeCount++;
            }
        }else{
            $nextSibling = $startParentNode->nextSibling;
            $nextNodeCount++;
            while(!is_null($nextSibling) && $this->getCommentRangeEnd($nextSibling,$id) === null && $nextSibling !== $commentRangeEndItem) {
                $nextSibling = $nextSibling->nextSibling;
                $nextNodeCount++;
            }
        }
        
        return ['parentNodeCount'=>$parentNodeCount,'nextNodeCount'=>$nextNodeCount];
    }
    
    private function getCommentRangeEnd($parentNode,$id) {
        $commentRangeEndItems = $parentNode->getElementsByTagName('commentRangeEnd');
        
        foreach($commentRangeEndItems as $commentRangeEndItem) {
            $eid = $this->getAttr($commentRangeEndItem, 'id');
            
            if($id === $eid) {
                return $commentRangeEndItem;
            }
        }
        
        return null;
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
        
        for($i=0;$i<$count;$i++) {
            $copy = clone $copyP;
            $this->insertAfter($copy, $p);
            $p = $copy;
        }
        
        return $p;
    }
    
    private function updateMDWORD_LINK($beginNode,$endNode,$link) {
        $mixed = [
            'r'=>[
                'childs'=>[
                    'fldChar'=>[
                        'fldCharType'=>'begin',
                        //                                 'text'=>'12',
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
