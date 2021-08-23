<?php
namespace MDword\Read;

use MDword\Edit\Part\Document;
use MDword\Edit\Part\Comments;
use MDword\Read\Part\ContentTypes;
use MDword\Common\Log;
use MDword\WordProcessor;
use MDword\Edit\Part\Styles;
use MDword\Edit\Part\Footer;
use MDword\Edit\Part\Header;

class Word 
{
    public $zip = null;
    /**
     * @var Log
     */
    public $log = null;
    
    public $Content_Types = null;
    
    private $tempDocumentFilename = null;
    /**
     * @var Header
     */
    public $headerEdits = [];
    /**
     * @var Document
     */
    public $documentEdit = null;
    /**
     * @var Footer
     */
    public $footerEdits = [];
    /**
     * @var Comments
     */
    public $commentsEdit = null;
    /**
     * @var Styles
     */
    public $stylesEdit = null;
    
    public $parts = [];
    
    public $files = [];
    
    public $blocks = [];
    
    /**
     * @var array
     */
    public $needUpdateParts = []; 
    
    /**
     * @var WordProcessor
     */
    public $wordProcessor = null;
    
    
    public function __construct() {
        $this->log = new Log();
    }
    
    
    public function load($archive) {
        $this->tempDocumentFilename = tempnam($this->getTempDir(), 'MDword');
        if (false === $this->tempDocumentFilename) {
            throw new \Exception('temp path make failed!');
        }
        
        
        if (false === file_put_contents($this->tempDocumentFilename, file_get_contents($archive)) || !is_file($this->tempDocumentFilename)) {
            throw new \Exception($archive.' copy file failed!');
        }
        
        $this->zip = new \ZipArchive();
        $this->zip->open($this->tempDocumentFilename);
        
        $this->read();
        
        //update needUpdateParts
        if(isset($this->parts[22]) && isset($this->blocks[22])) {
            foreach($this->blocks[22] as $partName => $value) {
                $this->needUpdateParts[$partName] = ['func'=>'getHeaderEdit','partName'=>$partName];
            }
        }
        
        $this->needUpdateParts['word/document.xml'] = ['func'=>'getDocumentEdit','partName'=>'word/document.xml'];
        
        if(isset($this->parts[23])  && isset($this->blocks[23])) {
            foreach($this->blocks[23] as $partName => $value) {
                $this->needUpdateParts[$partName] = ['func'=>'getFooterEdit','partName'=>$partName];
            }
        }
    }
    
    private function read() {
        $this->Content_Types = new ContentTypes($this->getXmlDom('[Content_Types].xml'));
        $this->Content_Types->word = $this;
        foreach ($this->Content_Types->overrides as $part) {
            if(in_array($part['ContentType'], [2,22,23])) {//document footer header
                $standardXmlFunc = function($xml) use($part) {
                    $xml = $this->standardXml($xml,$part['ContentType'],$part['PartName']);
                    return $xml;
                };
            }else{
                $standardXmlFunc = null;
            }
            
            if($part['ContentType'] === 14) {//image/png
                $this->parts[$part['ContentType']][] = ['PartName'=>$part['PartName'],'DOMElement'=>$part['PartName']];
            }else{
                $this->parts[$part['ContentType']][] = ['PartName'=>$part['PartName'],'DOMElement'=>$this->getXmlDom($part['PartName'],$standardXmlFunc)];
            }
        }
    }
    
    public static function getTempDir()
    {
        $tempDir = sys_get_temp_dir();
        
        if (!empty(self::$tempDir)) {
            $tempDir = self::$tempDir;
        }
        
        return $tempDir;
    }
    
    public function save($remainComments = false)
    {
        $this->deleteComments($remainComments);
        //update Toc
        $this->documentEdit->updateToc();
        
        $this->deleteComments($remainComments);
        
        foreach($this->parts as $type => $list ) {
            foreach($list as $part) {
                if(is_object($part['DOMElement'])) {
                    //delete space before break page
                    if($type === 2) {
                        $this->zip->addFromString($part['PartName'], $this->autoDeleteSpaceBeforeBreakPage($part['DOMElement']->saveXML()));
                    }else{
                        $this->zip->addFromString($part['PartName'], $part['DOMElement']->saveXML());
                    }
                }
            }
        }
        
        foreach($this->files as $part) {
            if(isset($part['this'])) {
                $this->zip->addFromString($part['PartName'], $part['this']->getContent());
            }else{
                $this->zip->deleteName($part['PartName']);
            }
        }
        
        $this->zip->addFromString('[Content_Types].xml', $this->Content_Types->DOMDocument->saveXML());
        
        if (false === $this->zip->close()) {
            throw new \Exception('Could not close zip file.');
        }
        
//         trigger_error('debug',E_USER_ERROR);exit;
        
        if(MDWORD_DEBUG === true) {
            $this->zip->open($this->tempDocumentFilename);
            $this->zip->extractTo(MDWORD_GENERATED_DIRECTORY);
            $this->zip->close();
        }
        
        return $this->tempDocumentFilename;
    }
    
    public function saveForTrace()
    {
        foreach($this->parts as $type => $list ) {
            foreach($list as $part) {
                if(is_object($part['DOMElement'])) {
                    //delete document end space
                    if($type === 2) {
                        $this->zip->addFromString($part['PartName'], $this->autoDeleteSpaceBeforeBreakPage($part['DOMElement']->saveXML()));
                    }else{
                        $this->zip->addFromString($part['PartName'], $part['DOMElement']->saveXML());
                    }
                }
            }
        }
        
        foreach($this->files as $part) {
            if(isset($part['this'])) {
                $this->zip->addFromString($part['PartName'], $part['this']->getContent());
            }else{
                $this->zip->deleteName($part['PartName']);
            }
        }
        
        $this->zip->addFromString('[Content_Types].xml', $this->Content_Types->DOMDocument->saveXML());
        
        
        $this->zip->open($this->tempDocumentFilename);
        $this->zip->extractTo(MDWORD_GENERATED_DIRECTORY);
        
        return $this->tempDocumentFilename;
    }
    
    
    private function standardXml($xml,$ContentType,$PartName) {
        $xml = preg_replace_callback('/\$[^$]*?\{[\s\S]+?\}/i', function($match){
            return preg_replace('/\s/', '', strip_tags($match[0]));
        }, $xml);
        
        static $commentId = 0;
        $nameToCommendId = [];
        $xml = preg_replace_callback('/(<[w|m]\:r[> ](?:(?!<[w|m]:r[> ])[\S\s])*?<[w|m]\:t[ ]{0,1}[^>]*?>)([^><]*?)(<\/[w|m]\:t>[\s\S]*?<\/[w|m]\:r>)/i', function($matchs) use(&$commentId,&$nameToCommendId,$ContentType,$PartName){
            return preg_replace_callback('/\$\{([\s\S]+?)\}/i', function($match) use(&$commentId,&$nameToCommendId,$matchs,$ContentType,$PartName){
                $name = $match[1];
                $length = strlen($name);
                if($name[$length-1] === '/') {
                    $name = trim($name,'/');
                    $this->blocks[$ContentType][$PartName]['r'.$commentId] = $name;
                    return $matchs[3].'<w:commentRangeStart w:id="r'.$commentId.'"/>'.$matchs[1].$match[0].$matchs[3].'<w:commentRangeEnd w:id="r'.$commentId++.'"/>'.$matchs[1];
                //end
                }elseif($name[0] === '/') {
                    $name = trim($name,'/');
                    return $match[0].$matchs[3].'<w:commentRangeEnd w:id="r'.$nameToCommendId[$name].'"/>'.$matchs[1];
                //start
                }else{
                    $name = trim($name,'/');
                    $this->blocks[$ContentType][$PartName]['r'.$commentId] = $name;
                    $nameToCommendId[$name] = $commentId;
                    return $matchs[3].'<w:commentRangeStart w:id="r'.$commentId++.'"/>'.$matchs[1].$match[0];
                }
                
            }, $matchs[0]);
        }, $xml);
        
        return $xml;
    }
    
    private function autoDeleteSpaceBeforeBreakPage($xml) {
        ini_set("pcre.backtrack_limit",-1);//回溯bug fixed,-1代表不做限制
        $xml = preg_replace_callback('/\<\/w\:p\>(?:\<w\:p\>(?:(?!\<\/w\:t\>|\<v\:imagedata).)+?\<\/w\:p\>|\<w\:p\/\>)*?\<w\:p[^>]*?\>\<w\:r\>\<w\:br w\:type\=\"page\"\/\>\<\/w\:r\>\<\/w\:p\>/i', function($match) {
            
            preg_match_all('/\<w\:p[^>]*?\>(.+?)\<\/w\:p\>/i', $match[0],$subMatch);
            if(!empty($subMatch[1])) {
                return implode('', $subMatch[1]).'</w:p>';
            }
            
            return '<w:r><w:br w:type="page"/></w:r></w:p>';
        }, $xml);
            
        $xml = preg_replace('/\<w\:p\/\>\<w\:sectPr\>/','<w:sectPr>',$xml);
        
        return $xml;
    }
    
    private function deleteComments($remainComments = true) {
        if($remainComments) {
            $edit = $this->commentsEdit[0];
            $willDeleted = [];
            $usedCommentIds = $this->wordProcessor->getUsedCommentIds();
            $edit->treeToListCallback($edit->DOMDocument,function($node) use($edit,$usedCommentIds,&$willDeleted) {
                if($node->localName == 'comment' && isset($usedCommentIds[$edit->getAttr($node,'id')])) {
                    $willDeleted[] = $node;
                }else{
                    return $node;
                }
            });

            foreach($willDeleted as $node) {
                $edit->removeChild($node);
            }
        }else{
            $parts = [];
            if(isset($this->parts[15])) {
                $parts = array_merge($parts,$this->parts[15]);
                unset($this->parts[15]);
            }
            if(isset($this->parts[16])) {
                $parts = array_merge($parts,$this->parts[16]);
                unset($this->parts[16]);
            }
            if(isset($this->parts[17])) {
                $parts = array_merge($parts,$this->parts[17]);
                unset($this->parts[17]);
            }
            foreach($parts as $part) {
                $this->zip->deleteName($part['PartName']);
            }
        }

        foreach($this->needUpdateParts as $part) {
            $func = $part['func'];
            switch ($func) {
                case 'getHeaderEdit':
                    $this->deleteMded($this->headerEdits[$part['partName']],$remainComments);
                    break;
                case 'getDocumentEdit':
                    $this->deleteMded($this->wordProcessor->getDocumentEdit(),$remainComments);
                    break;
                case 'getFooterEdit':
                    $this->deleteMded($this->footerEdits[$part['partName']],$remainComments);
                    break;
            }
        }
        
        //test
//         echo $this->documentEdit->DOMDocument->saveXML();exit;
    }
    
    private function deleteMded($edit,$remainComments=false) {
        $deleteTags = [
            'commentRangeStart'=>1,
            'commentRangeEnd'=>1,
            'commentReference'=>1
        ];
        
        $willDeleted = [];
        $usedCommentIds = $this->wordProcessor->getUsedCommentIds();
        if($remainComments) {
            $edit->treeToListCallback($edit->DOMDocument,function($node) use($edit,$deleteTags,$usedCommentIds,&$willDeleted) {
                if($edit->getAttr($node,'md',null)) {
                    $willDeleted[] = $node;
                }else{
                    $id = $edit->getAttr($node,'id');
                    if(isset($deleteTags[$node->localName]) && $id[0] !== 'r' && isset($usedCommentIds[$id])) {
                        $willDeleted[] = $node;
                    }else{
                        return $node;
                    }
                }
            });
        }else{
            $edit->treeToListCallback($edit->DOMDocument,function($node) use($edit,$deleteTags,$usedCommentIds,&$willDeleted) {
                if($edit->getAttr($node,'md',null) || isset($deleteTags[$node->localName])) {
                    $willDeleted[] = $node;
                }else{
                    return $node;
                }
            });
        }
            
        foreach($willDeleted as $node) {
            $edit->removeChild($node);
        }
    }
    
    /**
     *
     * @param string $filename
     * @return \DOMDocument
     */
    public function getXmlDom($filename,$standardXmlFunc=null) {
        $xml = $this->zip->getFromName($filename);
        if(!is_null($standardXmlFunc)) {
            $xml = $standardXmlFunc($xml);
        }
        $domDocument = new \DOMDocument();
        $domDocument->loadXML($xml);
        if(MDWORD_DEBUG === true) {
            $domDocument->formatOutput = true;
        }
        
        return $domDocument;
    }
    
    protected function getZipFiles() {
        static $pathIndx = null;
        
        if(is_null($pathIndx)) {
            for($i=0;$i < $this->zip->numFiles; $i++) {
                $name = $this->zip->getNameIndex($i);
                $nameArr = explode('/', $name);
                
                $item = &$pathIndx;
                foreach ($nameArr as $pathOrName) {
                    if(!isset($item[$pathOrName])) {
                        $item[$pathOrName] = [];
                    }
                    $item = &$item[$pathOrName];
                }
            }
            unset($item);
        }
        
        return $pathIndx;
    }
   
    public function getChartParts(){
        if(isset($this->parts[13])){
            $chartPartsArr[13] = $this->getPartsDetail($this->parts[13]);
        }
        if(isset($this->parts[25])){
            $chartPartsArr[25] = $this->parts[25];
        }
        if(isset($this->parts[26])){
            $chartPartsArr[26] = $this->parts[26];
        }
        return $chartPartsArr;
    }

    public function getPartsDetail($parts){
        foreach($parts as &$part){
            $partInfo = pathinfo($part['PartName']);
            $partNameRel = $partInfo['dirname'].'/_rels/'.$partInfo['basename'].'.rels';
            $chartRelDom = $this->getXmlDom($partNameRel);
            if(!empty($chartRelDom)){
                $part['chartRelDom'] = $chartRelDom;         
            }
            $Relationships = $chartRelDom->getElementsByTagName('Relationship');
            foreach($Relationships as $Relationship){
                $target = $Relationship->getAttribute('Target');
                if(strpos($target,'xlsx') != false){
                    $part['embeddings']['xml'] = $this->zip->getFromName(str_replace('../','word/',$target));
                    preg_match('/embeddings\/([\s\S]+)/',$target,$match);
                    $part['embeddings']['name'] = $match[1];
                }
            }
        }
        return $parts;
    }

    public function setChartParts($chartparts){
        foreach($chartparts as $key=>&$part){
            if($key==13){
                if(!empty($this->parts[13])){
                    foreach($this->parts[13] as $keypart){
                        $partName = $keypart['PartName'];
                        preg_match('/(\d+)/',$partName,$match);
                        $chartKey = $match[1];
                    }
                    foreach($part as &$p){
                        $chartKey++;
                        preg_match('/(\d+)/',$p['PartName'],$partmatch);
                        $p['PartName'] = str_replace($partmatch[1],$chartKey,$p['PartName']);
                        $styleKey[$partmatch[1]] = $chartKey;              
                    }
                    foreach($chartparts[25] as &$stylepart){
                        preg_match('/(\d+)/',$stylepart['PartName'],$styleMatch);
                        $stylepart['PartName'] = str_replace($styleMatch[1],$styleKey[$styleMatch[1]],$stylepart['PartName']);
                    }   
                    foreach($chartparts[26] as &$colorepart){
                        preg_match('/(\d+)/',$colorepart['PartName'],$colorMatch);
                        $colorepart['PartName'] = str_replace($colorMatch[1],$styleKey[$colorMatch[1]],$colorepart['PartName']);
                    }   
                    $this->parts[13] = array_merge($this->parts[13],$chartparts[13]);
                }else{
                    $this->parts[13] = $chartparts[13];
                }
            }else{
                $this->parts[$key] = empty($this->parts[$key])?$chartparts[$key]:array_merge($this->parts[$key],$chartparts[$key]);
            }
        }
    }

    public function getContentTypes(){
        $contentTypes = $this->Content_Types->overrides;

        return $contentTypes;
    }

    public function getChartEmbeddings(){
        $fileList = $this->getZipFiles();
        $relArr = $fileList['word']['charts']['_rels'];
        $embeddingsList = $fileList['word']['embeddings'];
        //$this->documentEdit->setChartRelValue($relArr);
        foreach($embeddingsList as $embeddingKey=>$val){
            $embeddingsXml[$embeddingKey] = $this->zip->getFromName('word/embeddings/'.$embeddingKey);
        }
        return $embeddingsXml;
    }

    public function setContentTypes(){
        foreach($this->parts[13] as $part){
            $partArr['PartName'] = $part['PartName'];
            $partArr['ContentType'] = 13;
            if(array_search($partArr,$this->Content_Types->overrides) === false){
                $newOverrides[] = $partArr;
            }
        }
        foreach($this->parts[25] as $part){
            $partArr['PartName'] = $part['PartName'];
            $partArr['ContentType'] = 25;
            if(array_search($partArr,$this->Content_Types->overrides) === false){
                $newOverrides[] = $partArr;
            }
        }
        foreach($this->parts[26] as $part){
            $partArr['PartName'] = $part['PartName'];
            $partArr['ContentType'] = 26;
            if(array_search($partArr,$this->Content_Types->overrides) === false){
                $newOverrides[] = $partArr;
            }
        }
        $this->Content_Types->setContent_types($newOverrides);
        $this->writeContentTypeXml();

    }

    public function writeContentTypeXml(){
        $Relationships = $this->Content_Types->DOMDocument->getElementsByTagName('Override');
        foreach($this->Content_Types->overrides as $key=>$overrides){
            if(!isset($Relationships[$key])){
                $copy = clone $Relationships[0];
                $copy->setAttribute('PartName','/'.$overrides['PartName']);
                $copy->setAttribute('ContentType',$this->Content_Types->contentTypes[$overrides['ContentType']]);
                $Relationships[0]->parentNode->appendChild($copy);
            }
        }
        $this->zip->addFromString('[Content_Types].xml', $this->Content_Types->DOMDocument->saveXML());
    }

    
    public function updateChartRel(){
        foreach($this->parts[13] as $part){
            if(isset($part['chartRelDom'])){
                $Relationships = $part['chartRelDom']->getElementsByTagName('Relationship');
                preg_match('/chart(\d+)/',$part['PartName'],$match);
                foreach($Relationships as $Relationship){
                    $target = $Relationship->getAttribute('Target');
                    $Relationship->setAttribute('Target',preg_replace('/(\d+)/',$match[1],$target));
                }
                $relArr['PartName'] =  $part['PartName'];
                $relArr['relName'] =  'word/charts/_rels/'.$match[0].'.xml.rels';
                $relArr['dom'] =  $part['chartRelDom'];

                $this->documentEdit->setChartRel($relArr);
            }
        }
    }

    public function free() {
        unset($this->Content_Types->word);
        unset($this->Content_Types);

        unset($this->zip,$this->log);
        $this->documentEdit->free();
        unset($this->documentEdit->DOMDocument);

        unset($this->documentEdit);
        unset($this->documentEdit->commentsEdit);
        unset($this->commentsEdit);
        unset($this->stylesEdit);

    }
}
