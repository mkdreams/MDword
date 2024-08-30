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
    private $outlineLvl = [1,3];

    public $levels = [];
    
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
        static $stylesEdit = null;
        if(is_null($stylesEdit)) {
            $stylesEdit = $this->word->wordProcessor->getStylesEdit();
        }

        if(isset($this->DOMDocument->documentElement->tagList['w:sdtContent'])) {
            $sdtContent = $this->DOMDocument->documentElement->tagList['w:sdtContent'][0];
        }else{
            $sdtContent = $this->DOMDocument;
        }

        //Compatible with pws

        if(isset($sdtContent->tagList['w:instrText'])) {
            $instrTexts = $sdtContent->tagList['w:instrText'];
        }else{
            $instrTexts = [];
        }
        $this->hyperlinkParentNodeArr = [];
        foreach($instrTexts as $instrText) {
            $text = $instrText->nodeValue;
            preg_match('/ PAGEREF ([\S]+?) \\\\h /i',$text,$match);
            if(isset($match[1]) && !isset($this->hyperlinkParentNodeArr[$match[1]])) {
                $this->hyperlinkParentNodeArr[$match[1]] = $this->getParentToNode($instrText);
            }
            preg_match('/ HYPERLINK \\\\l ([\S]+?) /i',$text,$match);
            if(isset($match[1])) {
                $match[1] = trim($match[1],'"');
            }
            if(isset($match[1]) && !isset($this->hyperlinkParentNodeArr[$match[1]])) {
                $this->hyperlinkParentNodeArr[$match[1]] = $this->getParentToNode($instrText);
            }

            preg_match('/TOC[\s\S]+?(\d+)\-(\d+)/i',$text,$match);
            if(isset($match[1])) {
                $this->outlineLvl[0] = intval($match[1])-1;
                $this->outlineLvl[1] = intval($match[2])-1;
            }
        }

        //fixed: some TOC jump not include:PAGEREF OR HYPERLINK
        if(isset($sdtContent->tagList['w:hyperlink'])) {
            $hyperlinks = $sdtContent->tagList['w:hyperlink'];
            foreach($hyperlinks as $hyperlink) {
                $parentNode = $hyperlink->parentNode;
                $anchor = $this->getAttr($hyperlink, 'anchor');
                if(!isset($this->hyperlinkParentNodeArr[$anchor])) {
                    $this->hyperlinkParentNodeArr[$anchor] = $parentNode;
                }
            }
        }

        if(isset($this->DOMDocument->documentElement->tagList['w:pPr'])) {
            $pPrs = $this->DOMDocument->documentElement->tagList['w:pPr'];
        }else{
            $pPrs = [];
        }
        $pPrLen = count($pPrs);
        for($i = 0;$i<$pPrLen;$i++) {
            $pPr = $pPrs[$i];
            $outlineLvl = $this->getVal($pPr,'outlineLvl');
            if(is_null($outlineLvl)) {
                $pStyleId = $this->getVal($pPr,'pStyle');
                if(is_null($pStyleId)) {
                    continue;
                }
                
                $pStyle = $stylesEdit->getStyleById($pStyleId);
                $pPrTemp = $pStyle->getElementsByTagName('pPr')->item(0);
                if(is_null($pPrTemp)) {
                    continue;
                }
                $outlineLvl = $this->getVal($pPrTemp,'outlineLvl');
            }

            if(!is_null($outlineLvl)) {
                $val = $outlineLvl;
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
                        if($copy->getElementsByTagName('hyperlink')->item(0)) {
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
                        }else{
                            $subBegainEnd = 0;
                            $fldChars = $copy->getElementsByTagName('fldChar');
                            foreach($fldChars as $fldChar) {
                                $r = $fldChar->parentNode;
                                $fldCharType = $this->getAttr($fldChar,'fldCharType');
                                if($fldCharType === 'begin') {
                                    $subBegainEnd++;
                                }elseif($fldCharType === 'end') {
                                    $subBegainEnd--;
                                }
                                
                                if($subBegainEnd < 0) {
                                    $this->markDelete($r);
                                }
                            }
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
        $updateCount = 0;

        if((strlen($name) === 32 && preg_match('/^[a-z0-9]+?$/iu',$name) === 1)|| is_array($name)) {//media md5
            $blocks = [null];
        }else{
            $blocks = $this->getBlocks($name);
            if(empty($blocks)) {
                $this->word->log->writeLog('not find name! name: '.$name);
            }
        }
        
        foreach($blocks as $key => $block) {
            $updateCount++;

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

        return $updateCount;
    }
    
    public function updateToc() {
        $titles = $this->getTitles();
        if(empty($titles)) {
            return ;
        }

        if(isset($this->DOMDocument->documentElement->tagList['w:sdtContent'])) {
            $sdtContent = $this->DOMDocument->documentElement->tagList['w:sdtContent'][0];
        }else{
            return ;
        }

        foreach($titles as $index => $title) {
            $anchor = $title['anchor'][0]['name'];
            $copy = $title['copy'];
            $ts = $copy->getElementsByTagName('t');
            $tLen = $ts->length;

            $numTIdx = 0;
            $numT = $ts->item($numTIdx);
            if(!is_null($numT) && $title['numText'] != '') {
                $numT->nodeValue = $title['numText'];
            }else{
                $numTIdx = -1;
            }
            
            $tIdx = $tLen-2;
            $t = $ts->item($tIdx);
            if(is_null($t)) {
                $tIdx = 0;
                $t = $ts->item($tIdx);
            }
            $this->setAttr($t, 'space', 'preserve','xml');
            $t->nodeValue  = $this->htmlspecialcharsBase($title['text']);

            for($i=0;$i<$tLen-1;$i++) {
                if($tIdx === $i || $numTIdx === $i) {
                    continue;
                }
                $r = $this->getParentToNode($ts->item($i),'r');
                $this->markDelete($r);
            }
            
            $pageT = $ts->item($tLen-1);
            if(!is_null($pageT) && $tLen-1 > 0) {
                $pageT->nodeValue = '';
            }

            $hyperlink_r = $this->getParentToNode($ts->item(0),'r');

            $isHyperlink = true;
            if($hyperlink = $copy->getElementsByTagName('hyperlink')->item(0)) {
                $this->setAttr($hyperlink, 'anchor', $title['anchor'][0]['name']);
                $textNextNode = $hyperlink;
            }else{
                $textNextNode = $hyperlink_r->nextSibling;
                $isHyperlink = false;
            }
            
            //TOC jump in the browser or app. Such as: whatsapp
            $hyperlink_temp = $this->createNodeByXml('<w:hyperlink w:anchor="'.$title['anchor'][0]['name'].'" w:history="1"></w:hyperlink>');
            $hyperlink_temp->appendChild($hyperlink_r);

            $instrTexts = $copy->getElementsByTagName('instrText');
            $instrTextToc = 'TOC \o "1-3" \h \z \u';
            foreach($instrTexts as $instrText) {
                if(strpos($instrText->nodeValue, 'PAGEREF') !== false) {
                    $instrText->nodeValue = " PAGEREF $anchor \h ";
                }elseif(strpos($instrText->nodeValue, 'HYPERLINK') !== false){
                    $instrText->nodeValue = ' HYPERLINK \l '.$anchor.' ';
                }elseif(strpos($instrText->nodeValue, 'TOC') !== false){
                    $instrTextToc = trim($instrText->nodeValue);

                    //index eq 0 not delete
                    if($index !== 0) {
                        $r = $instrText->parentNode;
                        $this->markDelete($r->previousSibling);
                        $this->markDelete($r);
                        $this->markDelete($r->nextSibling);
                    }
                }
            }
            
            if($index === 0 && $isHyperlink === true) {//first add specail node
                $fldCharBegin = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="begin" /></w:r>');
                $this->insertBefore($fldCharBegin, $textNextNode);
                
                
                $fldCharPreserve = $this->createNodeByXml('<w:r><w:instrText xml:space="preserve"> '.$instrTextToc.' </w:instrText></w:r>');
                $this->insertBefore($fldCharPreserve, $textNextNode);
                
                $fldCharSeparate = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="separate"/></w:r>');
                $this->insertBefore($fldCharSeparate, $textNextNode);
            }

            $this->insertBefore($hyperlink_temp,$textNextNode);

            $sdtContent->insertBefore($copy,$sdtContent->lastChild);
        }

        $fldChars = $sdtContent->getElementsByTagName('fldChar');
        $fldCharsCount = [];
        foreach($fldChars as $fldChar) {
            $fldCharType = $this->getAttr($fldChar, 'fldCharType');
            if(empty($this->getAttr($fldChar->parentNode, 'md',null))) {
                if(isset($fldCharsCount[$fldCharType])) {
                    $fldCharsCount[$fldCharType]++;
                }else{
                    $fldCharsCount[$fldCharType] = 1;
                }
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

        if(count($this->levels) === 0) {
            $this->levels = [];
        }

        $this->treeToListCallback($this->DOMDocument,function($node) use(&$titles) {
            if($node->localName === 'pPr') {
                $pPr = $node;

                $outlineLvl = $this->getVal($pPr,'outlineLvl');
                $pStyleId = $this->getVal($pPr,'pStyle');
                
                if(is_null($outlineLvl)) {
                    if(is_null($pStyleId)) {
                        return null;
                    }

                    $pStyle = $this->word->wordProcessor->getStylesEdit()->getStyleById($pStyleId);
                    $pPrTemp = $pStyle->getElementsByTagName('pPr')->item(0);
                    if(is_null($pPrTemp)) {
                        return null;
                    }
                    $outlineLvl = $this->getVal($pPrTemp,'outlineLvl');
                }

                if(!is_null($outlineLvl) && isset($this->anchors[$outlineLvl])
                                                    && $outlineLvl >= $this->outlineLvl[0]
                                                    && $outlineLvl <= $this->outlineLvl[1]
                                                ) {
                    $text = trim($this->getTextContent($pPr->parentNode));
                    if($text === '') {
                        return null;
                    }

                    $numText = $this->getNumText($pPr,$pStyleId);

                    $anchorInfo = [];
                    $anchorInfo = $this->updateBookMark($pPr->parentNode);

                    $anchorNode = $this->anchors[$outlineLvl];
                    $copy = clone $anchorNode;
    
                    $titles[] = ['copy'=>$copy,'anchor'=>$anchorInfo,'text'=>$text,'numText'=>$numText];

                    if($pStyleId > 0) {
                        $this->levels[] = ['index'=>count($this->levels),'level'=>$pStyleId,'name'=>$anchorInfo[0]['name'],'text'=>$this->getTextContent($pPr->parentNode)];
                    }
                }elseif($pStyleId > 0) {
                    $anchorInfo = [];
                    $anchorInfo = $this->updateBookMark($pPr->parentNode);

                    $this->levels[] = ['index'=>count($this->levels),'level'=>$pStyleId,'name'=>$anchorInfo[0]['name'],'text'=>$this->getTextContent($pPr->parentNode)];
                }
                return null;
            }else{
                return $node;
            }
        });

        return $titles;
    }

    private function getNumText($pPr,$id) {
        if(is_null($id)) {
            return '';
        }

        static $stylesEdit = null;
        static $numberingEdit = null;
        if(is_null($stylesEdit)) {
            $stylesEdit = $this->word->wordProcessor->getStylesEdit();
            $numberingEdit = $this->word->wordProcessor->getNumberingEdit();
        }

        $numPr = $pPr->getElementsByTagName('numPr')->item(0);
        if(is_null($numPr)) {
            $style = $stylesEdit->getStyleById($id);
            $numPr = $style->getElementsByTagName('numPr')->item(0);
        }

        if(is_null($numPr)) {
            return '';
        }

        $ilvl = $this->getVal($numPr,'ilvl');
        if(is_null($ilvl)) {
            $ilvl = 0;
        }

        return $numberingEdit->getText(intval($this->getVal($numPr,'numId')),$ilvl);
    }


    public function getLevels() {
        if(count($this->levels) === 0) {
            $this->getTitles();
        }
        return $this->levels;
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
            case MDWORD_TABLE:    
                $targetNode = $this->getTarget($nodeIdxs,'tcPr');
                $styles[$stylekey] = $targetNode;
                return $targetNode;
        }
    }

    private function getPStyle($name='',$type=MDWORD_TEXT) {
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
                $r = $this->getTarget($nodeIdxs,'r');
                if(is_null($r)) {
                    return null;
                }
                if($pPrs = $r->parentNode->getElementsByTagName('pPr')) {
                    $pPr = $pPrs->item(0);
                }
                $styles[$stylekey] = $pPr;
                return $styles[$stylekey];
                break;
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
                                    if(!isset($valueArr['text'])) {
                                        $valueArr['text'] = 1;
                                    }
                                    $this->updateMDWORD_BREAK_PAGE($targetNode->parentNode,$valueArr['text'],true);
                                    break;
                                case MDWORD_LINK:
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

                                    if(!is_null($valueArr['text'])) {
                                        $t = $copy->getElementsByTagName('t')->item(0);$this->setAttr($t, 'space', 'preserve','xml');$t->nodeValue = $this->htmlspecialcharsBase($valueArr['text']);
                                    }

                                    $this->insertBefore($copy, $targetNode);
                                    $this->updateMDWORD_LINK($copy, $copy, $valueArr['link']);
                                    if(!isset($value[$index+1])) {
                                        $this->markDelete($targetNode);
                                    }
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
                                    if(isset($valueArr['width'])) { 
                                        $imageInfo[1] = intval($imageInfo[1]*($valueArr['width']/$imageInfo[0]));
                                        $imageInfo[0] = $valueArr['width'];
                                    }else{
                                        //max width 550 px
                                        if($imageInfo[0] > 550){
                                            $imageInfo[1] = intval($imageInfo[1]*(550/$imageInfo[0]));
                                            $imageInfo[0] = 550;
                                        }
                                    }
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
                                    
                                    $copyP = $this->updateMDWORD_BREAK($targetNode->parentNode,1,false,true);
                                    $targetNode = $copyP->getElementsByTagName('r')->item(0);
                                    $this->removeMarkDelete($targetNode);
                                    if(!isset($value[$index+1])) {
                                        $this->markDelete($copyP);
                                    }
                                    break;
                                default:
                                    if(isset($valueArr['pstyle'])) {
                                        $pPr = $this->getPStyle($valueArr['pstyle']);
                                    }else{
                                        $pPr = null;
                                    }
                                    if(!is_null($pPr)) {
                                        $pPrCopy = clone $pPr;
                                        $pPrOrg = $targetNode->parentNode->getElementsByTagName('pPr')->item(0);
                                        if(is_null($pPrOrg)) {// pPr insert before t
                                            $this->insertBefore($pPrCopy, $targetNode->parentNode);
                                        }else{
                                            $this->insertBefore($pPrCopy, $targetNode->parentNode);
                                            $this->markDelete($pPrOrg);
                                        }
                                    }

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

                                    if(isset($valueArr['table_style'])){
                                        if(is_array($valueArr['table_style'])) {
                                            //[name=>'table_style','vMerge'=>2,'gridSpan'=>2]
                                            if(isset($valueArr['table_style']['name'])) {
                                                $trStyle = $this->getStyle($valueArr['table_style']['name'],MDWORD_TABLE);
                                            }else{
                                                $trStyle = null;
                                            }

                                            $vMergeCount = intval($valueArr['table_style']['vMerge']);
                                            $gridSpanCount = intval($valueArr['table_style']['gridSpan']);

                                            $tc = $this->getParentToNode($targetNode,'tc');
                                            if($vMergeCount > 1) {
                                                $trNextSibling = $tc->parentNode;
                                                $vMergeCount--;
                                                $vMergeNextCount = 0;
                                                $trNextSiblingNeedUpdateEnds = [$trNextSibling];
                                                while($vMergeCount > 0) {
                                                    $oldtrNextSibling = $trNextSibling;
                                                    $trNextSibling = $trNextSibling->nextSibling;
                                                    if(is_null($trNextSibling)) {
                                                        $trNextSibling = $oldtrNextSibling;
                                                        break;
                                                    }
                                                    $trNextSiblingNeedUpdateEnds[] = $trNextSibling;
                                                    $vMergeNextCount++;
                                                    $vMergeCount--;
                                                }


                                                if($vMergeNextCount > 0) {
                                                    //tc index
                                                    $tcIndex = $this->getIndex($tc->parentNode,$tc);

                                                    if($gridSpanCount === 0) {
                                                        $gridSpanCount = 1;
                                                    }
                                                    foreach($trNextSiblingNeedUpdateEnds as $idx => $trNextSibling) {
                                                        if($idx === 0) {
                                                            $vMergeXml = '<w:vMerge w:val="restart"/>';
                                                        }else{
                                                            $vMergeXml = '<w:vMerge/>';
                                                        }
                                                        if($gridSpanCount > 1) {
                                                            $gridSpanXml = '<w:gridSpan w:val="'.$gridSpanCount.'"/>';
                                                        }else{
                                                            $gridSpanXml = '';
                                                        }

                                                        $endTc = $trNextSibling->getElementsByTagName('tc')->item($tcIndex);
                                                        $tcPr = $endTc->getElementsByTagName('tcPr')->item(0);
                                                        if(!is_null($tcPr)) {
                                                            if($gridSpanXml !== '') {
                                                                $gridSpan = $tcPr->getElementsByTagName('gridSpan')->item(0);
                                                                if(is_null($gridSpan)) {
                                                                    $this->appendChild($tcPr,$this->createNodeByXml($gridSpanXml));
                                                                }
                                                            }

                                                            $vMerge = $tcPr->getElementsByTagName('vMerge')->item(0);
                                                            if(is_null($vMerge)) {
                                                                $this->appendChild($tcPr,$this->createNodeByXml($vMergeXml));
                                                            }

                                                        }else{
                                                            $tcPr = $this->createNodeByXml('<w:tcPr>'.$gridSpanXml.$vMergeXml.'</w:tcPr>');
                                                            if(is_null($endTc->firstChild)) {
                                                                $this->appendChild($endTc,$tcPr);
                                                            }else{
                                                                $this->insertBefore($tcPr,$endTc->firstChild);
                                                            }
                                                        }

                                                        $tcIndexTotal = $tcIndex+$gridSpanCount;
                                                        for($i=$tcIndex;$i<$tcIndexTotal;$i++) {
                                                            if($i === $tcIndex && $gridSpanCount > 1) {
                                                                $tcNextSibling = $endTc;
                                                                $gridSpanCountTemp = $gridSpanCount;
                                                                while((--$gridSpanCountTemp) > 0) {
                                                                    $tcNextSibling = $tcNextSibling->nextSibling;
                                                                    if(is_null($tcNextSibling)) {
                                                                        break;
                                                                    }
                
                                                                    $this->markDelete($tcNextSibling);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }

                                            }else if($gridSpanCount > 1) {
                                                if(is_null($trStyle)) {
                                                    $tcPr = $tc->getElementsByTagName('tcPr')->item(0);
                                                }else{
                                                    $tcPr = $trStyle;
                                                }
                                                if(!is_null($tcPr)) {
                                                    $gridSpan = $tcPr->getElementsByTagName('gridSpan')->item(0);
                                                    if(is_null($gridSpan)) {
                                                        $this->appendChild($tcPr,$this->createNodeByXml('<w:gridSpan w:val="'.$valueArr['table_style']['gridSpan'].'"/>'));
                                                    }
                                                }else{
                                                    $tcPr = $this->createNodeByXml('<w:tcPr><w:gridSpan w:val="'.$valueArr['table_style']['gridSpan'].'"/></w:tcPr>');
                                                }

                                                $tcNextSibling = $tc;
                                                while((--$gridSpanCount) > 0) {
                                                    $tcNextSibling = $tcNextSibling->nextSibling;
                                                    if(is_null($tcNextSibling)) {
                                                        break;
                                                    }

                                                    $this->markDelete($tcNextSibling);
                                                }
                                            }
                                        }else{
                                            $trStyle = $this->getStyle($valueArr['table_style'],MDWORD_TABLE);
                                        }
                                    }else{
                                        $trStyle = null;
                                    }

                                    if(!is_null($trStyle)){
                                        $trCopy = clone $trStyle;
                                        $tcPrStyle = $targetNode->parentNode->parentNode->getElementsByTagName('tcPr')->item(0);
                                        if(is_null($tcPrStyle)){
                                            $this->insertBefore($trCopy, $targetNode->parentNode);
                                        }else{
                                            $this->insertBefore($trCopy, $targetNode->parentNode);
                                            $this->markDelete($tcPrStyle);
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
                    if(empty($rids)) {
                        $this->word->log->writeLog('not find image by md5! md5: '.$name);
                    }

                    $rids = $name;
                    $handleFlag = false;
                    $domTmp = $this->domList;
                    $node = array_shift($domTmp);
                    $drawings = $node->getElementsByTagName('drawing');
                    $drawingArr = [];
                    foreach($drawings as $drawing) {
                        $blip = $drawing->getElementsByTagName('blip')->item(0);
                        if ($blip){
                            $rid = $this->getAttr($blip, 'embed', 'r');
                            if (in_array($rid, $rids)){
                                $drawingArr[$rid] = $drawing;
                            }
                        }
                    }

                    $imageInfo = @getimagesize($value);
                    foreach ($drawingArr as $rid => $drawing) {
                        $this->updateRef($value,$rid);
                        $extents = $drawing->getElementsByTagName('extent');
                        foreach($extents as $extent) {
                            $cx = $this->getAttr($extent, 'cx', null);
                            if($cx > 0) {
                                break;
                            }
                        }

                        if($cx <= 0 && $spPr = $drawing->getElementsByTagName('spPr')->item(0)) {
                            $exts = $spPr->getElementsByTagName('ext');
                            foreach($exts as $extent) {
                                $cx = $this->getAttr($extent, 'cx', null);
                                if($cx > 0) {
                                    break;
                                }
                            }
                        }

                        if($cx > 0) {
                            $orgCy = intval($cx/($imageInfo[0]*9530)*($imageInfo[1]*9530));
                            $extents = $drawing->getElementsByTagName('extent');
                            foreach($extents as $extent) {
                                $this->setAttr($extent, 'cy', $orgCy, null);
                            }
                            
                            if($spPr = $drawing->getElementsByTagName('spPr')->item(0)) {
                                $exts = $spPr->getElementsByTagName('ext');
                                foreach($exts as $extent) {
                                    $this->setAttr($extent, 'cy', $orgCy, null);
                                }
                            }
                        }
                    }
                    break;
                }

                $drawing = $this->getTarget($nodeIdxs,'drawing');
                if(!is_null($drawing)) {
                    $refInfo = $this->updateRef($value,null,MDWORD_IMG);
                    $rId = $refInfo['rId'];
                    $blip = $drawing->getElementsByTagName('blip')->item(0);
                    $this->setAttr($blip, 'embed', $rId ,'r');

                    //update cy
                    $extents = $drawing->getElementsByTagName('extent');
                    foreach($extents as $extent) {
                        $cx = $this->getAttr($extent, 'cx', null);
                        if($cx > 0) {
                            break;
                        }
                    }
                    
                    if($cx <= 0 && $spPr = $drawing->getElementsByTagName('spPr')->item(0)) {
                        $exts = $spPr->getElementsByTagName('ext');
                        foreach($exts as $extent) {
                            $cx = $this->getAttr($extent, 'cx', null);
                            if($cx > 0) {
                                break;
                            }
                        }
                    }

                    if($cx > 0) {
                        $orgCy = intval($cx/($refInfo['imageInfo'][0]*9530)*($refInfo['imageInfo'][1]*9530));
                        $extents = $drawing->getElementsByTagName('extent');
                        foreach($extents as $extent) {
                            $this->setAttr($extent, 'cy', $orgCy, null);
                        }
                        
                        if($spPr = $drawing->getElementsByTagName('spPr')->item(0)) {
                            $exts = $spPr->getElementsByTagName('ext');
                            foreach($exts as $extent) {
                                $this->setAttr($extent, 'cy', $orgCy, null);
                            }
                        }
                    }
                    break;
                }
                $drawing = $this->createNodeByXml('image');
                
                $refInfo = $this->updateRef($value,null,MDWORD_IMG);
                $rId = $refInfo['rId'];
                $imageInfo = $refInfo['imageInfo'];

                //max width 800
                if($imageInfo[1] > 800){
                    $imageInfo[0] = intval($imageInfo[0]*(800/$imageInfo[1]));
                    $imageInfo[1] = 800;
                }
                
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
            case MDWORD_CLONESECTION:
                if(is_array($value)) {
                    $value = $value[0];
                    $nameTo = $value[1];
                }else{
                    $nameTo = '';
                }

                list($sectionIdx,$nodeIdxs) = $this->getSectionNodeIdxs($nodeIdxs[0]);
                $lastNodeIdx = end($nodeIdxs);
                $isCloneLast = false;
                //last section
                if($this->domList[$sectionIdx]->localName === 'sectPr') {
                    $isCloneLast = true;
                    $sectPrXml = $this->DOMDocument->saveXML($this->domList[$sectionIdx]);
                }

                for($i=1;$i<$value;$i++) {
                    foreach($nodeIdxs as $nodeIdx) {
                        //section
                        if($sectionIdx === $nodeIdx) {
                            if($isCloneLast === true) {
                                if($value-1 === $i) {
                                    $prpr = $this->createNodeByXml($sectPrXml);
                                    $this->treeToList($prpr);

                                    //change last org section
                                    $this->markDelete($this->domList[$sectionIdx]);
                                    $orgSection = $this->createNodeByXml('<w:p><w:pPr>'.$sectPrXml.'</w:pPr></w:p>');
                                    $this->treeToList($orgSection);
                                    $this->insertBefore($orgSection,$this->domList[$sectionIdx]);
                                }else{
                                    $prpr = $this->createNodeByXml('<w:p><w:pPr>'.$sectPrXml.'</w:pPr></w:p>');
                                    $this->treeToList($prpr);
                                }
                                $this->insertAfter($prpr,$this->domList[$lastNodeIdx]);
                                $lastNodeIdx = $prpr->idxBegin;
                            }else{
                                $lastNodeIdx = $this->cloneNode($nodeIdx,$lastNodeIdx,$name,$i);
                            }
                            $headerReference = $this->domList[$lastNodeIdx]->getElementsByTagName('headerReference')->item(0);
                            if(!is_null($headerReference)) {
                                $rid = $this->getAttr($headerReference,'id','r');
                                $cloneInfo = $this->getRels()->cloneRels($rid);
                                $this->setAttr($headerReference,'id',$cloneInfo['rId'],'r');

                                $this->word->needUpdateParts[$cloneInfo['partName']] = [
                                    'func'=>'getHeaderEdit',
                                    'partName'=>$cloneInfo['partName'],
                                ];

                                $standardXmlFunc = function() use($cloneInfo) {
                                    $xml = $this->word->standardXml($cloneInfo['xml'],22,$cloneInfo['partName']);
                                    return $xml;
                                };
                                $this->word->parts[22][] = ['PartName'=>$cloneInfo['partName'],'DOMElement'=>$this->word->getXmlDom(null,$standardXmlFunc)];
                                $headerEdit = $this->word->wordProcessor->getHeaderEdit($cloneInfo['partName']);
                                foreach($headerEdit->domIdxToName as $nodeIdxTemp => $v) {
                                    $headerEdit->cloneNode($nodeIdxTemp,$nodeIdxTemp,$name,0,$i);
                                }
                            }

                            $footerReference = $this->domList[$lastNodeIdx]->getElementsByTagName('footerReference')->item(0);
                            if(!is_null($footerReference)) {
                                $rid = $this->getAttr($footerReference,'id','r');
                                $cloneInfo = $this->getRels()->cloneRels($rid);
                                $this->setAttr($footerReference,'id',$cloneInfo['rId'],'r');

                                $this->word->needUpdateParts[$cloneInfo['partName']] = [
                                    'func'=>'getFooterEdit',
                                    'partName'=>$cloneInfo['partName'],
                                ];

                                $standardXmlFunc = function() use($cloneInfo) {
                                    $xml = $this->word->standardXml($cloneInfo['xml'],23,$cloneInfo['partName']);
                                    return $xml;
                                };
                                $this->word->parts[23][] = ['PartName'=>$cloneInfo['partName'],'DOMElement'=>$this->word->getXmlDom(null,$standardXmlFunc)];
                                $footerEdit = $this->word->wordProcessor->getFooterEdit($cloneInfo['partName']);
                                foreach($footerEdit->domIdxToName as $nodeIdxTemp => $v) {
                                    $footerEdit->cloneNode($nodeIdxTemp,$nodeIdxTemp,$name,0,$i);
                                }
                            }
                        }else{
                            $lastNodeIdx = $this->cloneNode($nodeIdx,$lastNodeIdx,$name,$i);
                        }
                    }
                }
                
                //刷新被克隆对象
                foreach($nodeIdxs as $nodeIdx) {
                    $lastNodeIdx = $this->cloneNode($nodeIdx,$lastNodeIdx,$name,0);
                    //section
                    if($sectionIdx === $nodeIdx) {
                        $headerReference = $this->domList[$lastNodeIdx]->getElementsByTagName('headerReference')->item(0);
                        if(!is_null($headerReference)) {
                            $rid = $this->getAttr($headerReference,'id','r');
                            $Relationship = $this->getRels()->ridToTarget[$rid];
                            $type = $this->getAttr($Relationship, 'Type');
                            $oldTarget = $Relationship->getAttribute('Target');
                            $headerEdit = $this->word->wordProcessor->getHeaderEdit('word/'.$oldTarget);
                            foreach($headerEdit->domIdxToName as $nodeIdxTemp => $v) {
                                $headerEdit->cloneNode($nodeIdxTemp,$nodeIdxTemp,$name,0);
                            }
                        }

                        $footerReference = $this->domList[$lastNodeIdx]->getElementsByTagName('footerReference')->item(0);
                        if(!is_null($footerReference)) {
                            $rid = $this->getAttr($footerReference,'id','r');
                            $Relationship = $this->getRels()->ridToTarget[$rid];
                            $type = $this->getAttr($Relationship, 'Type');
                            $oldTarget = $Relationship->getAttribute('Target');
                            $footerEdit = $this->word->wordProcessor->getFooterEdit('word/'.$oldTarget);
                            foreach($footerEdit->domIdxToName as $nodeIdxTemp => $v) {
                                $footerEdit->cloneNode($nodeIdxTemp,$nodeIdxTemp,$name,0);
                            }
                        }
                    }
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
                }elseif($value == 'tr') {
                    foreach($nodeIdxs as $nodeIdx) {
                        $tr = $this->getParentToNode($nodeIdx,'tr');
                        if(!is_null($tr)) {
                            $this->markDelete($tr);
                        }
                    }
                }elseif($value == 'section') {
                    list($sectionIdx,$nodeIdxs) = $this->getSectionNodeIdxs($nodeIdxs[0]);
                    foreach($nodeIdxs as $nodeIdx) {
                        $this->markDelete($this->domList[$nodeIdx]);
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
                $this->updateMDWORD_LINK($nodeIdxs[0], $nodeIdxs[strlen($nodeIdxs)-1], $value);
                break;
            case MDWORD_REF:
                $r = $this->getTarget($nodeIdxs,'r',function($node) {
                    $t = $node->getElementsByTagName('t');
                    if($t->length > 0) {
                        return true;
                    }else{
                        return false;
                    }
                });
                $this->updateMDWORD_REF($r,$value);
                break;
            case MDWORD_PAGEREF:
                $r = $this->getTarget($nodeIdxs,'r',function($node) {
                    $t = $node->getElementsByTagName('t');
                    if($t->length > 0) {
                        return true;
                    }else{
                        return false;
                    }
                });
                $this->updateMDWORD_PAGEREF($r,$value);
                break;
            case MDWORD_NOWPAGE:
                $r = $this->getTarget($nodeIdxs,'r',function($node) {
                    $t = $node->getElementsByTagName('t');
                    if($t->length > 0) {
                        return true;
                    }else{
                        return false;
                    }
                });
                $this->updateMDWORD_PRESERVE($r,isset($value['text'])?$value['text']:null,'PAGE');
                break;
            case MDWORD_TOTALPAGE:
                $r = $this->getTarget($nodeIdxs,'r',function($node) {
                    $t = $node->getElementsByTagName('t');
                    if($t->length > 0) {
                        return true;
                    }else{
                        return false;
                    }
                });
                $this->updateMDWORD_PRESERVE($r,isset($value['text'])?$value['text']:null,'NUMPAGES');
                break;
            case MDWORD_PHPWORD:
                //get p
                $p = $this->getParentToNode($nodeIdxs[0]);
                    
                $XmlFromPhpword = new XmlFromPhpword($value,$this);
                $nodes = $XmlFromPhpword->createNodesByBodyXml();
                
                foreach($nodes as $node) {
                    $copy = clone $node;
                    $this->insertBefore($copy, $p);
                }
                
                $this->markDelete($p);
                break;
            default:
                break;
        }
    }
    
    private function cloneNode($nodeIdx,$endNodeIdx,$name,$idx,$reIdx=0) {
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
                        $newName = $nameTemp[1].'#'.$reIdx;
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
            $name = $this->getAttr($bookmarkStart,'name');
            if(strpos($name,'_MD') !== 0) {
                $maxId = $maxId+1;
                $name = '_MD_Toc'.$maxId;
                $this->setAttr($bookmarkStart, 'id', $maxId);
                $this->setAttr($bookmarkStart, 'name', $name);
                $id = $maxId;
            }else{
                $id = $this->getAttr($bookmarkStart,'id');
            }
            $infos[] = ['id'=>$id,'name'=>$name];

            $ids[$key] = $id;
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
                $name = '_MD_Toc'.$maxId;
                $bookmarkStart = $this->createNodeByXml('<w:bookmarkStart w:id="'.$maxId.'" w:name="'.$name.'"/>');
                $this->insertBefore($bookmarkStart, $rs->item(0));
                
                $bookmarkEnd = $this->createNodeByXml('<w:bookmarkEnd w:id="'.$maxId.'"/>');
                $this->insertAfter($bookmarkEnd, $rs->item($rs->length-1));
                
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
        if(is_int($beginNodeIndex)) {
            $parentNode = $this->domList[$beginNodeIndex];
        }else{
            $parentNode = $beginNodeIndex;
        }
        while($parentNode->localName != $type && !is_null($parentNode)) {
            $parentNode = $parentNode->parentNode;
        }
        
        return $parentNode;
    }

    private function getSectionNodeIdxs($beginNodeIndex) {
        static $cache = [];

        if(isset($cache[$beginNodeIndex])) {
            return $cache[$beginNodeIndex];
        }

        $nodeP = $this->getParentToNode($beginNodeIndex);

        $nodeIdxs = [$nodeP->idxBegin];
        $previousSibling = $nextSibling = $nodeP;

        $preIdxs = [];
        while($previousSibling->previousSibling) {
            $sectPr = $previousSibling->previousSibling->getElementsByTagName('sectPr');
            if($sectPr->length > 0) {
                break;
            }
            $previousSibling = $previousSibling->previousSibling;
            $preIdxs[] = $previousSibling->idxBegin;
        }

        $endIdxs = [];
        while($nextSibling->nextSibling) {
            $nextSibling = $nextSibling->nextSibling;
            $endIdxs[] = $nextSibling->idxBegin;
            $sectPr = $nextSibling->getElementsByTagName('sectPr');
            if($sectPr->length > 0) {
                break;
            }
        }

        $nodeIdxs = array_merge(array_reverse($preIdxs),$nodeIdxs,$endIdxs);

        $cache[$beginNodeIndex] = [$nextSibling->idxBegin,$nodeIdxs];

        return [$nextSibling->idxBegin,$nodeIdxs];
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

    public function getRels() {
        if(is_null($this->rels)) {
            $this->initRels();
        }
        return $this->rels;
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
    
    

    private function updateMDWORD_PRESERVE($r,$text=null,$preserve='') {
        //fixed update fields style change
        $rPr = $r->getElementsByTagName('rPr')->item(0);
        if(!is_null($rPr)) {
            $rPrXml = $this->DOMDocument->saveXML($rPr);
        }else{
            $rPrXml = '';
        }

        if(is_null($text)) {
            $text = 'Please Update Field';
        }
        $begin = $this->createNodeByXml('<w:r>'.$rPrXml.'<w:fldChar w:fldCharType="begin"/></w:r>');

        $preserve = $this->createNodeByXml('<w:r>'.$rPrXml.'<w:instrText xml:space="preserve">'.$preserve.'</w:instrText></w:r>');
        
        $separate = $this->createNodeByXml('<w:r>'.$rPrXml.'<w:fldChar w:fldCharType="separate"/></w:r>');
        
        $end = $this->createNodeByXml('<w:r>'.$rPrXml.'<w:fldChar w:fldCharType="end"/></w:r>');
        
        $this->insertBefore($begin,$r);
        $this->insertBefore($preserve,$r);
        $this->insertBefore($separate,$r);
        $this->insertAfter($end,$r);

        $t = $r->getElementsByTagName('t')->item(0);
        $t->nodeValue = $this->htmlspecialcharsBase($text);
    }

    private function updateMDWORD_REF($r,$value,$type='REF') {
        $this->updateMDWORD_PRESERVE($r,isset($value['text'])?$value['text']:null,' '.$type.' '.$value['name'].' \h ');
    }

    private function updateMDWORD_PAGEREF($r,$value) {
        $this->updateMDWORD_REF($r,$value,'PAGEREF');
    }

    private function updateMDWORD_LINK($beginNode,$endNode,$link) {
        $link = $this->htmlspecialcharsBase($link,false);
        if(strpos($link,'#') === 0) {
            $link = ltrim($link,'#');
            $hyperlinkNodeBegin = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="begin"/></w:r>');
            $hyperlinkNodePreserve = $this->createNodeByXml('<w:r><w:instrText xml:space="preserve">HYPERLINK </w:instrText></w:r>');
            $hyperlinkNodePreserveTwo = $this->createNodeByXml('<w:r><w:instrText xml:space="preserve"> \l "'.$link.'" </w:instrText></w:r>');
            $hyperlinkNodeSeparate = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="separate"/></w:r>');
            $hyperlinkNodeEnd = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="end"/></w:r>');
            
            $this->insertBefore($hyperlinkNodeBegin, $beginNode);
            $this->insertBefore($hyperlinkNodePreserve, $beginNode);
            $this->insertBefore($hyperlinkNodePreserveTwo, $beginNode);
            $this->insertBefore($hyperlinkNodeSeparate, $beginNode);
            $this->insertAfter($hyperlinkNodeEnd, $endNode);
        }else{

            $hyperlinkNodeBegin = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="begin"/></w:r>');
            $hyperlinkNodePreserve = $this->createNodeByXml('<w:r><w:instrText xml:space="preserve"> HYPERLINK "'.$link.'" \o "'.$link.'" </w:instrText></w:r>');
            $hyperlinkNodeSeparate = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="separate"/></w:r>');
            $hyperlinkNodeEnd = $this->createNodeByXml('<w:r><w:fldChar w:fldCharType="end"/></w:r>');
            
            $this->insertBefore($hyperlinkNodeBegin, $beginNode);
            $this->insertBefore($hyperlinkNodePreserve, $beginNode);
            $this->insertBefore($hyperlinkNodeSeparate, $beginNode);
            $this->insertAfter($hyperlinkNodeEnd, $endNode);
        }
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
