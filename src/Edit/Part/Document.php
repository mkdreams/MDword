<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Document extends PartBase
{
    public $commentsblocks;
    public function __construct($word,\DOMDocument $DOMDocument,$blocks = []) {
        parent::__construct($word);
        
        $this->DOMDocument = $DOMDocument;
        $this->commentsblocks = $blocks;
        
        $this->initNameSpaces();
    }
    
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
                $copy->getElementsByTagName('t')->item(0)->nodeValue= $title['text'];
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
    
    private function update($block,$value,$type) {
        $beginNode = $block[0];
        
        $endNode = $block[1];
        
        $traces = $block[2];
        $parentNodeCount = $traces['parentNodeCount'];
        $nextNodeCount = $traces['nextNodeCount'];
        switch ($type) {
            case 'text':
                $targetNode = $this->getTarget($beginNode,$endNode,$parentNodeCount);
                if(!is_null($targetNode)) {
                    $copy = clone $targetNode;
                    $copy->getElementsByTagName('t')->item(0)->nodeValue= $value;
                    $parentNode = $targetNode->parentNode;
                    $parentNode->insertBefore($copy,$targetNode); 
                    
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
                    $this->getExcelPath($rid);
                    var_dump($targetNode);exit;
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
        
//         $targetNode
        $this->removeMarkDelete($targetNode);
        return $targetNode;
    }
    
    private function getExcelPath($rid='') {
        if(is_null($this->rels)) {
            $this->getRels();
        }
        
        
        $Relationships = $this->rels->DOMDocument->getElementsByTagName('Relationship');
        $length = $Relationships->length;
        foreach ($Relationships as $Relationship) {
            if($Relationship->getAttribute('Id') === $rid) {
                $Target = $this->rels->partInfo['dirname'].'/'.$Relationship->getAttribute('Target');
                var_dump($Target);exit;
            }
        }
    }
    
    private function updateRef($rid,$file) {
        if(is_null($this->rels)) {
            $this->getRels();
        }
        
        $Relationships = $this->rels->DOMDocument->getElementsByTagName('Relationship');
        $length = $Relationships->length;
        foreach ($Relationships as $Relationship) {
            if($Relationship->getAttribute('Id') === $rid) {
                $oldValue = $this->rels->partInfo['dirname'].'/'.$Relationship->getAttribute('Target');
                $target = 'media/image'.++$length.'.png';
                $Relationship->setAttribute('Target',$target);
                $target = $this->rels->partInfo['dirname'].'/'.$target;
                $this->word->zip->addFromString($target, file_get_contents($file));
                $this->word->zip->deleteName($oldValue);
            }
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
