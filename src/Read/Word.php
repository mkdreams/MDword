<?php
namespace MDword\Read;

use MDword\Edit\Part\Document;
use MDword\Edit\Part\Comments;
use MDword\Read\Part\ContentTypes;
use MDword\Common\Log;
use MDword\WordProcessor;

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
     * @var Document
     */
    public $documentEdit = null;
    /**
     * @var Comments
     */
    public $commentsEdit = null;
    
    public $parts = [];
    
    public $files = [];
    
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
            throw new \Exception('temp path make faild!');
        }
        
        if (false === copy($archive, $this->tempDocumentFilename)) {
            throw new \Exception($archive.'copy file fiald!');
        }
        
        
        $this->zip = new \ZipArchive();
        $this->zip->open($this->tempDocumentFilename);
        
        $this->read();
    }
    
    private function read() {
        $this->Content_Types = new ContentTypes($this->getXmlDom('[Content_Types].xml'));
        foreach ($this->Content_Types->overrides as $part) {
            if($part['ContentType'] === 14) {//image/png
                $this->parts[$part['ContentType']][] = ['PartName'=>$part['PartName'],'DOMElement'=>$part['PartName']];
            }else{
                $this->parts[$part['ContentType']][] = ['PartName'=>$part['PartName'],'DOMElement'=>$this->getXmlDom($part['PartName'])];
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
    
    public function save()
    {
        $this->deleteComments();
        
        //update Toc
        $this->documentEdit->updateToc();
        
        //delete again
        $this->deleteComments();
        
        foreach($this->parts as $type => $list ) {
            foreach($list as $part) {
                if(is_object($part['DOMElement'])) {
                    //delete document end space
                    if($type === 2) {
                        $this->zip->addFromString($part['PartName'], $this->autoDeleteSpacePage($part['DOMElement']->saveXML()));
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
        
        if(MDWORD_DEBUG === true) {
            $this->zip->open($this->tempDocumentFilename);
            $this->zip->extractTo(MDWORD_GENERATED_DIRECTORY);
            $this->zip->close();
        }
        
        return $this->tempDocumentFilename;
    }
    
    public function saveForTrace()
    {
        $this->deleteComments();
        
        //update Toc
        $this->documentEdit->updateToc();
        
        //delete again
        $this->deleteComments();
        
        foreach($this->parts as $type => $list ) {
            foreach($list as $part) {
                if(is_object($part['DOMElement'])) {
                    //delete document end space
                    if($type === 2) {
                        $this->zip->addFromString($part['PartName'], $this->autoDeleteSpacePage($part['DOMElement']->saveXML()));
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
    
    
    private function autoDeleteSpacePage($xml) {
        return $xml;
//         $xml = preg_replace_callback('/\<\/w\:p\>(?:\<w\:p\>(?:(?!\<\/w\:t\>).)+?\<\/w\:p\>|\<w\:p\/\>)*?\<w\:p.+?\<w\:r.+?\<w\:br w\:type\=\"page\"\/\>\<\/w\:r\>\<\/w\:p\>/i', function($match) {
//         preg_match('/\<w\:r\>(?:(?!\<\/w\:r\>).)*?\<w\:br w\:type\=\"page\"\/\>\<\/w\:r\>\<\/w\:p\>/i', $xml, $match);
//         var_dump($match);exit;
        $xml = preg_replace_callback('/<\/w:p>(?:<w:p[ >](?:(?!(<\/w:t>|)).)+?<\/w:p>|<w:p\/>)*?\<w\:p(?:(?!<\/w:p>).)*?<w:r>(?:(?!<w:r>).)*?<w:br w:type\=\"page\"\/><\/w:r><\/w:p>/i', function($match) {
//             echo $match[0];
//             var_dump($match);
            preg_match_all('/<w:p[ >](.+?)\<\/w:p>/i', $match[0],$subMatch);
            
            if(!empty($subMatch[1])) {
                return implode('', $subMatch[1]).'</w:p>';
            }
            
            return '<w:r><w:br w:type="page"/></w:r></w:p>';
        }, $xml);
//         echo $xml;exit;
        
        return $xml;
    }
    
    private function deleteComments() {
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

        
        //remove marked
        $this->documentEdit->deleteMarked();
        
        //test
//         echo $this->documentEdit->DOMDocument->saveXML();exit;
        
        //remove comments tag
        $this->documentEdit->deleteByXpath('//w:commentRangeStart|//w:commentRangeEnd|//w:commentReference/..');
    }
    
    /**
     *
     * @param string $filename
     * @return \DOMDocument
     */
    public function getXmlDom($filename) {
        $xml = $this->zip->getFromName($filename);
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
}
