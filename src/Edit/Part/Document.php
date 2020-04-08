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
//         var_dump($rid);exit;
        $partInfo = pathinfo($this->partName);
        if(is_null($this->refDOMDocument)) {
            $this->partNameRel = $partInfo['dirname'].'/_rels/'.$partInfo['basename'].'.rels';
            $this->refDOMDocument = $this->word->getXmlDom($this->partNameRel);
            $this->word->parts[19][] = ['PartName'=>$this->partNameRel,'DOMElement'=>$this->refDOMDocument];
        }
        
        
        $Relationships = $this->refDOMDocument->getElementsByTagName('Relationship');
        $length = $Relationships->length;
        foreach ($Relationships as $Relationship) {
            if($Relationship->getAttribute('Id') === $rid) {
                $Target = $partInfo['dirname'].'/'.$Relationship->getAttribute('Target');
                var_dump($Target);exit;
            }
        }
    }
    
    private function updateRef($rid,$file) {
        $partInfo = pathinfo($this->partName);
        if(is_null($this->refDOMDocument)) {
            $this->partNameRel = $partInfo['dirname'].'/_rels/'.$partInfo['basename'].'.rels';
            $this->refDOMDocument = $this->word->getXmlDom($this->partNameRel);
            $this->word->parts[19][] = ['PartName'=>$this->partNameRel,'DOMElement'=>$this->refDOMDocument];
        }
        
        
        $Relationships = $this->refDOMDocument->getElementsByTagName('Relationship');
        $length = $Relationships->length;
        foreach ($Relationships as $Relationship) {
            if($Relationship->getAttribute('Id') === $rid) {
                $oldValue = $partInfo['dirname'].'/'.$Relationship->getAttribute('Target');
                $target = 'media/image'.++$length.'.png';
                $Relationship->setAttribute('Target',$target);
                $target = $partInfo['dirname'].'/'.$target;
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
