<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Document extends PartBase
{
    public $commentsblocks;
    public $charts = [];
    
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
            'text'=>'value blue',
            'style'=>'blue'
            ],
            [
            'text'=>'value red',
            'style'=>'red'
            ],
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
        
        $hyperlinkArr = [];
        foreach($hyperlinks as $hyperlink) {
            $hyperlinkArr[$this->getAttr($hyperlink, 'anchor')] = $hyperlink;
        }
        
        $index = 0;
        foreach($titles as $anchor => $title) {
            $anchorOrg = $title['orgname'];
            
            if(isset($hyperlinkArr[$anchorOrg])) {
                $index++;
                $p = $hyperlinkArr[$anchorOrg]->parentNode;
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
        
        
        foreach($hyperlinkArr as $hyperlink) {
            $this->markDelete($hyperlink->parentNode);
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
    
    private function getStyle($name='') {
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
        
        $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,'r');
        if($rPrs = $targetNode->getElementsByTagName('rPr')) {
            $rPr = $rPrs->item(0);
        }
        
        return $rPr;
    }
    
    private function update($block,$value,$type) {
        $beginNode = $block[0];
        $endNode = $block[1];
        
        $traces = $block[2];
        $parentNodeCount = $traces['parentNodeCount'];
        $nextNodeCount = $traces['nextNodeCount'];
        switch ($type) {
            case 'text':
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,'r');
                if(!is_null($targetNode)) {
                    if(is_array($value)) {
                        foreach($value as $valueArr) {
                            $copy = clone $targetNode;
                            $rPr = $this->getStyle($valueArr['style']);
                            if(!is_null($rPr)) {
                                $rPrCopy = clone $rPr;
                                $rPrOrg = $copy->getElementsByTagName('rPr')->item(0);
                                
//                                 $rPrOrgParentNode = $rPrOrg->parentNode;
//                                 $rPrOrgParentNode->insertBefore($rPrCopy,$rPrOrg);
                                $this->insertBefore($rPrCopy, $rPrOrg);
                                
                                $this->markDelete($rPrOrg);
                            }
                            
                            $copy->getElementsByTagName('t')->item(0)->nodeValue= $valueArr['text'];
                            $this->insertBefore($copy, $targetNode);
//                             $parentNode = $targetNode->parentNode;
//                             $parentNode->insertBefore($copy,$targetNode);
                        }
                        
//                         var_dump($value);exit;
                    }else{
                        $copy = clone $targetNode;
                        $copy->getElementsByTagName('t')->item(0)->nodeValue= $value;
                        $this->insertBefore($copy, $targetNode);
//                         $parentNode = $targetNode->parentNode;
//                         $parentNode->insertBefore($copy,$targetNode); 
                    }
                    
                    $this->markDelete($targetNode);
                }
//                 $this->deleteMarked();
//                 echo $this->DOMDocument->saveXML();exit;
                break;
            case 'image':
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,'pict');
                if(!is_null($targetNode)) {
                    $rid = $this->getAttr($targetNode->getElementsByTagName('imagedata')->item(0), 'id', 'r');
                    $this->updateRef($rid,$value);
                }
                break;
            case 'excel':
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,'drawing');
                if(!is_null($targetNode)) {
                    $rid = $this->getAttr($targetNode->getElementsByTagName('chart')->item(0), 'id', 'r');
                    $this->initChart($rid);
                    $value = $this->charts[$rid]->excel->preDealDatas($value);
                    $this->charts[$rid]->excel->changeExcelValues($value);
                    $this->charts[$rid]->chartRelUpdateByType($value,'str');
                    $this->charts[$rid]->chartRelUpdateByType($value,'num');
                }
                break;
            case 'clone':
                $parentNode = $beginNode;
                $targetNode = null;
                $needCloneNodes = [];
                for($i=0;$i<$parentNodeCount;$i++) {
                    $parentNode = $parentNode->parentNode;
                }
                
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount,'r');
                
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
                        $this->updateCommentsId($copy, $i,$value);
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
                break;
            case 'delete':
                if($value == 'p') {
                    $p = $this->getParentToNode($beginNode,'p');
                }else{
                    
                }
                $this->markDelete($p);
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
            $this->setAttr($bookmarkStart, 'orgname', $this->getAttr($bookmarkStart, 'name'));
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
            $this->setAttr($commentRangeEnd, 'id', $name);
        }
    }
    
    private function getTarget($beginNode,$endNode,$parentNodeCount,$type='r') {
        $parentNode = $beginNode;
        $targetNode = null;
        for($i=0;$i<=$parentNodeCount;$i++) {
            $nextSibling = $parentNode;
            while($nextSibling = $nextSibling->nextSibling) {
                if($nextSibling === $endNode) {
                    break 2;
                }
                if(is_null($targetNode)) {
                    if($nextSibling->localName == $type) {//is target
                        $targetNode = $nextSibling;
                    }else{//sub find target
                        $rs = $nextSibling->getElementsByTagName($type);
                        if($rs->length > 0) {
                            $targetNode = $rs->item(0);
                        }else{
                            $this->markDelete($nextSibling);
                        }
                    }
                }else{
                    $this->markDelete($nextSibling);
                }
            }
            
            if($i === $parentNodeCount) {//top parent
                
            }else{
                $parentNode = $parentNode->parentNode;
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
    
    private function updateRef($rid,$file) {
        if(is_null($this->rels)) {
            $this->initRels();
        }
        
        $this->rels->replace($rid,$file);
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
        $commentRangeStartItems = $this->DOMDocument->getElementsByTagName('commentRangeStart');
        
        $blocks = [];
        foreach($commentRangeStartItems as $commentRangeStartItem) {
            $id = $this->getAttr($commentRangeStartItem, 'id');
            if($this->commentsblocks[$id] !== $name) {
                continue;
            }
            $commentRangeEndItem = $this->getCommentRangeEnd($this->DOMDocument,$id);
            
            $trace = $this->getRangeTrace($id,$commentRangeStartItem, $commentRangeEndItem);
            $blocks[] = [$commentRangeStartItem,$commentRangeEndItem,$trace];
        }
        
        return $blocks;
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
            
            $endParentNode = $commentRangeEndItem;
        }else{
            $nextSibling = $startParentNode->nextSibling;
            $nextNodeCount++;
            while(!is_null($nextSibling) && $this->getCommentRangeEnd($nextSibling,$id) === null && $nextSibling !== $commentRangeEndItem) {
                $nextSibling = $nextSibling->nextSibling;
                $nextNodeCount++;
            }
            $endParentNode = $nextSibling;
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
    
    
}
