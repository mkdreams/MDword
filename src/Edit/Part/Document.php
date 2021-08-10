<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;
use MDword\XmlTemple\XmlFromPhpword;

class Document extends PartBase
{
    public $commentsblocks;
    public $charts = [];
    public $blocks = [];
    public $usedBlock = [];
    private $anchors = [];
    private $hyperlinkParentNodeArr = [];
    
    public function __construct($word,\DOMDocument $DOMDocument,$blocks = []) {
        parent::__construct($word);
        
        $this->DOMDocument = $DOMDocument;
        $this->commentsblocks = $blocks;
        
        $this->treeToList($this->DOMDocument->documentElement);
        $this->initNameSpaces();
        $this->initLevelToAnchor();
        
        if(!$this->word->wordProcessor->isForTrace) {
            $this->blocks = $this->initCommentRange();
        }
    }
    
    private function initLevelToAnchor() {
        $sdtContent = $this->DOMDocument->documentElement->tagList['w:sdtContent'][0];
        if(is_null($sdtContent)) {
            $sdtContent = $this->DOMDocument;
        }

        $hyperlinks = $sdtContent->tagList['w:hyperlink'];
        $this->hyperlinkParentNodeArr = [];
        foreach($hyperlinks as $hyperlink) {
            $parentNode = $hyperlink->parentNode;
            $this->hyperlinkParentNodeArr[$this->getAttr($hyperlink, 'anchor')] = $parentNode;
        }
        $pPrs = $this->DOMDocument->documentElement->tagList['w:pPr'];
        $pPrLen = count($pPrs);
        for($i = 0;$i<$pPrLen;$i++) {
            $pPr = $pPrs[$i];
            $pStyle = $pPr->firstChild;
            if($pStyle->localName === 'pStyle') {
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
     * @param string|array $value examples：
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
    public function setValue($name,$value,$type=MDWORD_TEXT,$needRecord=true) {
        if(strlen($name) === 32 || is_array($name)) {//media md5
            $blocks = [null];
        }else{
            $blocks = $this->getBlocks($name);
            if(empty($blocks)) {
                $this->word->log->writeLog('not find name! name: '.$name);
            }
        }
        
        foreach($blocks as $key => $block) {
            $this->update($block,$name,$value,$type);
            //update node idx
            if(!is_null($block)) {
                $this->blocks[$name][$key] = $block;
            }
            
            if(!$needRecord) {
                continue;
            }
            //--SAVE-ANIMALCODE--
//--SAVE-ANIMALCODE--
        }
    }
    
    public function updateToc() {
        $titles = $this->getTitles();
        if(empty($titles)) {
            return ;
        }
        $sdtContent = $this->DOMDocument->documentElement->tagList['w:sdtContent'][0];
        if(is_null($sdtContent)) {
            return ;
        }
        
        foreach($titles as $index => $title) {
            $anchor = $title['anchor'][0]['name'];
            $copy = $title['copy'];
            $t = $copy->getElementsByTagName('t')->item(0);$this->setAttr($t, 'space', 'preserve','xml');$t->nodeValue  = $this->htmlspecialcharsBase($title['text']);
            
            if($tItme = $copy->getElementsByTagName('t')->item(1)) {
                $tItme->nodeValue = '';
            }
            $hyperlink = $copy->getElementsByTagName('hyperlink')->item(0);
            $this->setAttr($hyperlink, 'anchor', $title['anchor'][0]['name']);
            
            $instrTexts = $copy->getElementsByTagName('instrText');
            foreach($instrTexts as $instrText) {
                if(strpos($instrText->nodeValue, 'PAGEREF') > 0) {
                    $instrText->nodeValue = " PAGEREF $anchor \h ";
                }
            }
            
            if($index === 0) {//first add specail node
                $fldCharBegin = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="begin" /></w:r>');
                $this->insertBefore($fldCharBegin, $hyperlink);
                
                
                $fldCharPreserve = $this->createNodeByXml('<w:r><w:instrText xml:space="preserve"> TOC \o "1-3" \h \z \u </w:instrText></w:r>');
                $this->insertBefore($fldCharPreserve, $hyperlink);
                
                $fldCharSeparate = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="separate"/></w:r>');
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
            $fldCharEnd = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="end"/></w:r>');
            $copy->appendChild($fldCharEnd);
        }
        
        
        foreach($this->hyperlinkParentNodeArr as $hyperlinkParentNode) {
            $this->removeChild($hyperlinkParentNode);
        }
    }
    
    private function getTitles() {
        $titles = [];
        $this->treeToListCallback($this->DOMDocument,function($node) use(&$titles) {
            if($node->localName === 'pStyle') {
                $val = intval($this->getAttr($node, 'val'));
                if(isset($this->anchors[$val])) {
                    $pPr = $node->parentNode;
                    $anchorInfo = [];
                    $anchorInfo = $this->updateBookMark($pPr->parentNode);
                    $anchorNode = $this->anchors[$val];
                    $copy = clone $anchorNode;
    
                    $titles[] = ['copy'=>$copy,'anchor'=>$anchorInfo,'text'=>$pPr->parentNode->textContent];
                }
            }else{
                return $node;
            }
        });
        
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
        static $defaultStyle = null;
        
        $stylekey = $name.$type;
        
        if(isset($styles[$stylekey])) {
            return $styles[$stylekey];
        }
        
        if(is_null($defaultStyle)) {
            $defaultStyle = [];
            $stylesEdit = $this->word->wordProcessor->getStylesEdit();
            $tempStyles = $stylesEdit->DOMDocument->getElementsByTagName('style');
            foreach($tempStyles as $style) {
                if($this->getAttr($style, 'default') == 1 && $this->getAttr($style, 'type') == 'paragraph') {
                    if($rPr = $style->getElementsByTagName('rPr')) {
                        $rPr = $rPr->item(0);
                        $defaultStyle['rPr'] = $rPr;
                    }
                }
            }
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
                
                $styles[$stylekey] = $this->mergerPr($rPr, $defaultStyle['rPr']);
                return $rPr;
                break;
            case MDWORD_IMG:
                $targetNode = $this->getTarget($nodeIdxs,'drawing');
                $styles[$stylekey] = $targetNode;
                return $targetNode;
        }
    }
    
    public function update(&$nodeIdxs,$name,$value,$type) {
        switch ($type) {
            case MDWORD_TEXT:
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
                    foreach($value as $index => $valueArr) {
                        $copy = clone $targetNode;
                        if(is_array($valueArr)) {
                            switch ($valueArr['type']){
                                case MDWORD_BREAK:
                                    $valueArr['text'] = intval($valueArr['text']);
                                    
                                    if(isset($value[$index-1]) && $value[$index-1]['type'] === MDWORD_IMG) {
                                        $valueArr['text']--;
                                    }
                                    
                                    if($valueArr['text'] <= 0) {
                                        break;
                                    }
                                    
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
                                    $this->updateMDWORD_BREAK_PAGE($targetNode->parentNode,$valueArr['text'],true);
                                    break;
                                case MDWORD_LINK:
                                    if(!is_null($valueArr['text'])) {
                                        $t = $copy->getElementsByTagName('t')->item(0);$this->setAttr($t, 'space', 'preserve','xml');$t->nodeValue = $this->htmlspecialcharsBase($valueArr['text']);
                                    }
                                    $this->insertBefore($copy, $targetNode);
                                    $this->markDelete($targetNode);
                                    $this->updateMDWORD_LINK($copy, $copy, $valueArr['link']);
                                    break;
                                case MDWORD_IMG:
                                    $drawing = $this->getStyle($valueArr['style'],MDWORD_IMG);
                                    $pStyle = $this->getStyle($valueArr['style']);
                                    if(!is_null($pStyle)) {
                                        $pPrStyle = $pStyle->parentNode->parentNode->getElementsByTagName('pPr')->item(0);
                                        $pPrCopy = clone $pPrStyle;
                                        $this->insertBefore($pPrCopy, $targetNode);
                                    }

                                    if(is_null($drawing)) {
                                        $drawing = $this->createNodeByXml('image');
                                    }
                                    $copyDrawing = clone $drawing;
                                    
                                    $refInfo = $this->updateRef($valueArr['text'],null,MDWORD_IMG);
                                    $rId = $refInfo['rId'];
                                    $imageInfo = $refInfo['imageInfo'];
                                    
                                    $docPr = $copyDrawing->getElementsByTagName('docPr')->item(0);
                                    preg_match("(\d+)",$rId,$idMatch);
                                    $this->setAttr($docPr, 'id', $idMatch[0], null);

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
                                    
                                    if(isset($value[$index-1]) && $value[$index-1]['type'] !== MDWORD_BREAK) {
                                        $this->markDelete($targetNode);
                                        $copyP = $this->updateMDWORD_BREAK($targetNode->parentNode,1,false);
                                        $targetNode = $copyP->getElementsByTagName('r')->item(0);
                                        $this->removeMarkDelete($targetNode);
                                    }
                                    
                                    $targetNode->getElementsByTagName('t')->item(0)->nodeValue= '';
                                    $targetNode->appendChild($copyDrawing);
                                    
                                    $copyP = $this->updateMDWORD_BREAK($targetNode->parentNode,1,false,false);
                                    $targetNode = $copyP->getElementsByTagName('r')->item(0);
                                    $this->removeMarkDelete($targetNode);
                                    if(!isset($value[$index+1])) {
                                        $this->markDelete($copyP);
                                    }
                                    break;
                                default:
                                    if(isset($valueArr['style'])) {
                                        $rPr = $this->getStyle($valueArr['style']);
                                    }else{
                                        $rPr = null;
                                    }
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
                                    $t = $copy->getElementsByTagName('t')->item(0);$this->setAttr($t, 'space', 'preserve','xml');$t->nodeValue = $this->htmlspecialcharsBase($valueArr['text']);
                                    $this->insertBefore($copy, $targetNode);
                                    break;
                            }
                        }else{
                            $t = $copy->getElementsByTagName('t')->item(0);$this->setAttr($t, 'space', 'preserve','xml');$t->nodeValue = $this->htmlspecialcharsBase($valueArr);
                            $this->insertBefore($copy, $targetNode);
                        }
                        
                    }
                    
                }else{
                    $copy = clone $targetNode;
                    $t = $copy->getElementsByTagName('t')->item(0);$this->setAttr($t, 'space', 'preserve','xml');$t->nodeValue = $this->htmlspecialcharsBase($value);
                    $this->treeToList($copy);
                    $this->extendIds($targetNode->idxBegin,[$copy->idxBegin]);
                    $this->insertBefore($copy, $targetNode);
                }
                
                $this->markDelete($targetNode);
                break;
            case MDWORD_IMG:
                if(is_null($nodeIdxs)) {//md5
                    $rids = $name;
                    if(empty($rids)) {
                        $this->word->log->writeLog('not find image by md5! md5: '.$name);
                    }
                    foreach($rids as $rid) {
                        $this->updateRef($value,$rid);
                    }
                    break;
                }
                
                $drawing = $this->getTarget($nodeIdxs,'drawing');
                if(!is_null($drawing)) {
                    $refInfo = $this->updateRef($value,null,MDWORD_IMG);
                    $rId = $refInfo['rId'];
                    $blip = $drawing->getElementsByTagName('blip')->item(0);
                    $this->setAttr($blip, 'embed', $rId ,'r');
                    break;
                }
                
                $drawing = $this->createNodeByXml('image');
                
                $refInfo = $this->updateRef($value,null,MDWORD_IMG);
                $rId = $refInfo['rId'];
                $imageInfo = $refInfo['imageInfo'];
                
                $docPr = $drawing->getElementsByTagName('docPr')->item(0);
                preg_match("(\d+)",$rId,$idMatch);
                $this->setAttr($docPr, 'id', $idMatch[0], null);

                $blip = $drawing->getElementsByTagName('blip')->item(0);
                $this->setAttr($blip, 'embed', $rId,'r');
                
                $orgCx = intval($imageInfo[0]*9530);
                $orgCy = intval($imageInfo[1]*9530);
                
                $extents = $drawing->getElementsByTagName('extent');
                foreach($extents as $extent) {
                    $this->setAttr($extent, 'cx', $orgCx, null);
                    $this->setAttr($extent, 'cy', $orgCy, null);
                }
                
                if($spPr = $drawing->getElementsByTagName('spPr')->item(0)) {
                    $exts = $spPr->getElementsByTagName('ext');
                    foreach($exts as $extent) {
                        $this->setAttr($extent, 'cx', $orgCx, null);
                        $this->setAttr($extent, 'cy', $orgCy, null);
                    }
                }
                
                $targetNode = $this->getTarget($nodeIdxs,'r',function($node) {
                    $t = $node->getElementsByTagName('t');
                    if($t->length > 0) {
                        return true;
                    }else{
                        return false;
                    }
                });
                
                $t = $targetNode->getElementsByTagName('t')->item(0);
                $this->markDelete($t);
                $targetNode->appendChild($drawing);
                
                // $copyP = $this->updateMDWORD_BREAK($targetNode->parentNode,1,false);
                // $targetNode = $copyP->getElementsByTagName('r')->item(0);
                // $this->removeMarkDelete($targetNode);
                break;
            case MDWORD_CLONEP:
                $p = $this->getParentToNode($nodeIdxs[0],'p');
                $nodeIdxs = [$p->idxBegin];
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
            case MDWORD_CLONETO:
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
            case MDWORD_CLONE:
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
            case MDWORD_DELETE:
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
            case MDWORD_BREAK:
                $p = $lastNode = $this->getParentToNode($nodeIdxs[0],'p');
                $this->updateMDWORD_BREAK($p,$value,true);
                break;
            case MDWORD_PAGE_BREAK:
                $p = $this->getParentToNode($nodeIdxs[0],'p');
                $this->updateMDWORD_BREAK_PAGE($p,$value,true);
                break;
            case MDWORD_LINK:
                $this->updateMDWORD_LINK($beginNode, $endNode, $value);
                break;
            case MDWORD_PHPWORD:
                //get p
                $targetNode = $this->getParentToNode($nodeIdxs[0]);
                    
                $XmlFromPhpword = new XmlFromPhpword($value,$this);
                $nodes = $XmlFromPhpword->createNodesByBodyXml();
                
                foreach($nodes as $node) {
                    $copy = clone $node;
                    $this->insertBefore($copy, $targetNode);
                }
                
                $this->markDelete($targetNode);
                
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
            $beginIdOldToNew = $this->getTreeToListBeginIdOldToNew($node, $cloneNode->idxBegin);
            
            for($i = $begin; $i <= $end; $i++) {
                if(isset($this->domIdxToName[$i])) {
                    $cloneNodeIdx = $beginIdOldToNew[$i];
                    if(isset($this->idxExtendIdxs[$i])) {
                        $this->idxExtendIdxs[$cloneNodeIdx] = [];
                        foreach ($this->idxExtendIdxs[$i] as $oldId) {
                            $this->idxExtendIdxs[$cloneNodeIdx][] = $beginIdOldToNew[$oldId];
                        }
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
                $bookmarkStart = $this->createNodeByXml('<w:bookmarkStart w:id="'.$maxId.'" w:name="'.$name.'"/>');
                $this->insertBefore($bookmarkStart, $rs->item(0));
                
                $bookmarkEnd = $this->createNodeByXml('<w:bookmarkEnd w:id="'.$maxId.'"/>');
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
        $targetNode = null;
        foreach($nodeIdxs as $nodeIdx) {
            $node = $this->domList[$nodeIdx];
            $this->markDelete($node);
            if($find) {
                continue;
            }
            
            if($node->localName === $type && (is_null($checkCallBack) || $checkCallBack($node))) {
                $targetNode = $node;
                $find = true;
            }else{
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
    
    public function updateRef($file,$rid=null,$type=MDWORD_IMG) {
        if(is_null($this->rels)) {
            $this->initRels();
        }
        
        if(is_null($rid)) {
            return $this->rels->insert($file,$type);
        }else{
            return $this->rels->replace($rid,$file);
        }
    }
    
    public function getRidByMd5($md5) {
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

        //record use block
        $this->usedBlock[$name] = 1;
        
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
        
        //--HIGHLIGHT-ANIMALCODE--
        //--HIGHLIGHT-ANIMALCODE--
        
        return $blocks;
    }
    
    
    private function updateMDWORD_BREAK($p,$count=1,$replace=true,$isCopyP=true) {
        $copyP = $this->createNodeByXml('<w:p></w:p>');
        if($isCopyP){
            if($pPr = $p->getElementsByTagName('pPr')->item(0)) {
                $this->appendChild($copyP, clone $pPr);
            }
        }
        
        $t = $p->getElementsByTagName('t')->item(0);
        
        if($replace === true) {
            $this->markDelete($p);
        }
        
        $indexs = [];
        $copy = $copyP;
        for($i=0;$i<$count;$i++) {
            $baseIndex = $this->treeToList(null);
            $this->treeToList($copy);
            $indexs[] = $baseIndex;
            $this->insertAfter($copy, $p);
            $p = $copy;
            
            if($i < $count - 1) {
                $copy = clone $copyP;
            }
        }
        
        if(count($indexs) > 0) {
            $this->extendIds($copyP->idxBegin,$indexs);
        }
        
        $r = clone $t->parentNode;
        $t = $r->getElementsByTagName('t')->item(0)->nodeValue = ' ';
        foreach($r->getElementsByTagName('drawing') as $drawing) {
            $this->removeChild($drawing);
        }
        
        $this->appendChild($p, $r);
        
        return $p;
    }
    
    private function updateMDWORD_BREAK_PAGE($p,$count=1,$replace=true) {
        $copyP = $this->createNodeByXml('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
        if($replace === true) {
            $this->markDelete($p);
        }
        
        $indexs = [];
        $copy = $copyP;
        for($i=0;$i<$count;$i++) {
            $baseIndex = $this->treeToList(null);
            $this->treeToList($copy);
            $indexs[] = $baseIndex;
            $this->insertAfter($copy, $p);
            $p = $copy;
            
            if($i < $count - 1) {
                $copy = clone $copyP;
            }
        }
        
        if(count($indexs) > 0) {
            $this->extendIds($copyP->idxBegin,$indexs);
        }
        
        return $p;
    }
    
    private function updateMDWORD_LINK($beginNode,$endNode,$link) {
        $hyperlinkNodeBegin = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="begin"/></w:r>');
        $link = $this->htmlspecialcharsBase($link);
        $hyperlinkNodePreserve = $this->createNodeByXml('<w:r><w:instrText xml:space="preserve"> HYPERLINK "'.$link.'" \o "'.$link.'" </w:instrText></w:r>');
        
        $hyperlinkNodeSeparate = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="separate"/></w:r>');
        
        $hyperlinkNodeEnd = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="end"/></w:r>');
        
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
    /**
     * 
     * @param int $id
     * @param [] $newIds
     */
    private function extendIds($id,$newIds) {
        if(!isset($this->idxExtendIdxs[$id])) {
            $this->idxExtendIdxs[$id] = [];
        }
        $this->idxExtendIdxs[$id] = array_merge($this->idxExtendIdxs[$id],$newIds);
    }
    
    public function getBlockTree() {
        $domIdxToName = [];
        foreach($this->blocks as $name => $blocks) {
            foreach($blocks as $block) {
                $idxBegin = null;
                $idxEnd = null;
                
                foreach ($block as $index => $id) {
                    $node = $this->domList[$id];
                    if($this->isTc($node) && $index > 0) {
                        $node = $node->parentNode;
                    }
                    
                    if($node->idxBegin < $idxBegin || is_null($idxBegin)) {
                        $idxBegin = $node->idxBegin;
                    }
                    if($node->idxEnd > $idxEnd || is_null($idxEnd)) {
                        $idxEnd = $node->idxEnd;
                    }
                    
                    
                }
                $domIdxToName[$idxBegin] = [0,$name];
                $domIdxToName[$idxEnd] = [1,$name];
            }
        }
        ksort($domIdxToName);
        
        return $this->blockNameTree($domIdxToName);
    }
    
    private function blockNameTree($domIdxToName) {
        $tree = [];
        $name = null;
        foreach($domIdxToName as $key => $info) {
            if($info[0] === 0 && is_null($name)) {
                $name = $info[1];
                $tree[$name] = $key;
                $domIdxToNameTemp = [];
                continue;
            }
            
            if($info[0] === 1 && $name === $info[1]) {
                if(!empty($domIdxToNameTemp)) {
                    $tree[$name] = $this->blockNameTree($domIdxToNameTemp);
                }
                $name = null;
                continue;
            }
            
            $domIdxToNameTemp[$key] = $info;
        }
        
        return $tree;
    }
}
