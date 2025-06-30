<?php
namespace MDword;

use MDword\Edit\Part\Document;
use MDword\Read\Word;
use MDword\Edit\Part\Comments;
use MDword\Common\Bind;
use MDword\Edit\Part\Header;
use MDword\Edit\Part\Footer;
use MDword\Edit\Part\Styles;
use MDword\Edit\Part\Numbering;

class WordProcessor
{
    private $wordsIndex = -1;
    private $words = [];
    public $isForTrace = false;
    
    public function __construct() {
        error_reporting(E_ERROR | E_PARSE );
        require_once(dirname(__FILE__).DIRECTORY_SEPARATOR.'Config'.DIRECTORY_SEPARATOR.'Main.php');
    }
    
    public function load($zip) {
        $reader = new Word();
        $reader->wordProcessor = $this;
        $reader->load($zip);
        $this->words[++$this->wordsIndex] = $reader;
        
        $this->words[$this->wordsIndex]->commentsEdit = [];
        foreach ($this->words[$this->wordsIndex]->parts[15] as $part15) {
            $comments = $part15['DOMElement'];
            $Comment = new Comments($this->words[$this->wordsIndex],$comments);
            $Comment->partName = $part15['PartName'];
            $Comment->word = $this->words[$this->wordsIndex];
            $this->words[$this->wordsIndex]->commentsEdit[] = $Comment;
        }
        
        return $this->words[$this->wordsIndex];
    }
    
    /**
     * @return \MDword\Common\Bind
     */
    public function getBind($data) {
        $bind = new Bind($this,$data);
        return $bind;
    }
    
    public function setValue($name, $value, $type=MDWORD_TEXT) {
        $updateCount = 0;
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $updateCount += $documentEdit->setValue($name, $value, $type);
        }

        return $updateCount;
    }
    
    public function setValues($values,$pre='') {
        $updateCount = 0;
        foreach ($values as $index => $valueArr) {
            foreach($valueArr as $name => $value) {
                if(is_array($value)) {
                    $updateCount += $this->setValues($value,'#'.$index);
                }else{
                    $updateCount += $this->setValue($name.$pre.'#'.$index, $value);
                }
            }
        }
        return $updateCount;
    }

    public function getInnerVars() {
        return ['levels'=>$this->words[$this->wordsIndex]->documentEdit->getLevels()];
    }

    /**
     * delete section include the block。(only in document.xml)
     * @param string $name
     */
    public function deleteSection($name) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, 'section',MDWORD_DELETE);
        }
    }
    
    /**
     * delete p include the block 
     * @param string $name
     */
    public function deleteP($name) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, 'p',MDWORD_DELETE);
        }
    }
    
    /**
     * delete tbl include the block
     * @param string $name
     */
    public function deleteTbl($name) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, 'tbl',MDWORD_DELETE);
        }
    }

    /**
     * delete tr include the block
     * @param string $name
     */
    public function deleteTr($name) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, 'tr',MDWORD_DELETE);
        }
    }
    
    /**
     * delete block
     * @param string $name
     */
    public function delete($name) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, '',MDWORD_TEXT);
        }
    }
    
    public function setImageValue($name, $value) {
        if(strlen($name) === 32) {//media md5
            $includeImageEdits = [];
            
            $documentEdit = $this->getDocumentEdit();
            $rids = $documentEdit->getRidByMd5($name);
            if(!empty($rids)) {
                $includeImageEdits[] = [$documentEdit,$rids];
            }
            
            foreach ($this->words[$this->wordsIndex]->parts[22] as $part) {
                $partName = $part['PartName'];
                $headerEdit = $this->getHeaderEdit($partName);
                $rids = $headerEdit->getRidByMd5($name);
                if(!empty($rids)) {
                    $includeImageEdits[] = [$headerEdit,$rids];
                    $this->words[$this->wordsIndex]->needUpdateParts[$partName] = ['func'=>'getHeaderEdit','partName'=>$partName];
                }
            }
            
            foreach ($this->words[$this->wordsIndex]->parts[23] as $part) {
                $partName = $part['PartName'];
                $footerEdit = $this->getFooterEdit($partName);
                $rids = $footerEdit->getRidByMd5($name);
                if(!empty($rids)) {
                    $includeImageEdits[] = [$footerEdit,$rids];
                    $this->words[$this->wordsIndex]->needUpdateParts[$partName] = ['func'=>'getFooterEdit','partName'=>$partName];
                }
            }
            
            $edit = null;
            foreach($includeImageEdits as list($edit,$rids)) {
                /**
                 * @var Document $edit
                 */
                if (is_array($value)) {
                    $edit->setValue($rids, $value, MDWORD_TEXT);
                }else{
                    $edit->setValue($rids, $value, MDWORD_IMG);
                }
            }
            
            return ;
        }
        
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, $value,MDWORD_IMG);
        }
    }
    
    /**
     * @param string $name
     * ['text','link']
     * @param array $value
     */
    public function setLinkValue($name, $value) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, $value[0],MDWORD_TEXT);
            $documentEdit->setValue($name, $value[1],MDWORD_LINK);
        }
    }
    
//     /**
//      * @param string $name
//      * @param array $datas
//      * change value ['A1',9,'set']
//      * extention range ['$A$1:$A$5','$A$1:$A$10','ext']
//      */
//     public function setExcelValues($name='',$datas=[]) {
//         $documentEdit = $this->getDocumentEdit();
//         $documentEdit->setValue($name, $datas, 'excel');
//     }
    
    /**
     * clone p
     * @param string $name
     * @param int $count
     */
    public function cloneP($name,$count=1) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, $count, MDWORD_CLONEP);
        }
    }
    /**
     * clone section
     * @param string $name 
     * @param int|array $count 1|[1,$nameTo]
     */
    public function cloneSection($name,$count=1) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, $count, MDWORD_CLONESECTION);
        }
    }
    /**
     * clone
     * @param string $name
     * @param int $count
     */
    public function clones($name,$count=1) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, $count, MDWORD_CLONE);
        }
    }
    /**
     * clone
     * @param string $name
     * @param int $count
     */
    public function cloneTo($nameTo,$name) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($nameTo, $name, MDWORD_CLONETO);
        }
    }
    
    
    public function setBreakValue($name, $value) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, $value,MDWORD_BREAK);
        }
    }
    
    
    public function setBreakPageValue($name, $value=1) {
        foreach($this->words[$this->wordsIndex]->needUpdateParts as $part) {
            $func = $part['func'];
            $partName = $part['partName'];
            /**
             * @var Document $documentEdit
             */
            $documentEdit = $this->$func($partName);
            $documentEdit->setValue($name, $value,MDWORD_PAGE_BREAK);
        }
    }
    
    /**
     * update toc
     */
    public function updateToc() {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->updateToc();
    }
    
    /**
     * @var Document
     */
    public function getHeaderEdit($partName='word/header1.xml') {
        if(!isset($this->words[$this->wordsIndex]->headerEdits[$partName])) {
            $index = $this->getPartsIndexByPartName(22,$partName);
            $document = $this->words[$this->wordsIndex]->parts[22][$index]['DOMElement'];
            $blocks = $this->words[$this->wordsIndex]->blocks[22][$partName];//header not add comment
            $blocks = $blocks?$blocks:[];
            $headerEdit = new Header($this->words[$this->wordsIndex],$document,$blocks);
            $this->words[$this->wordsIndex]->headerEdits[$partName] = $headerEdit;
            $this->words[$this->wordsIndex]->headerEdits[$partName]->partName = $this->words[$this->wordsIndex]->parts[22][$index]['PartName'];
        }else{
            $headerEdit = $this->words[$this->wordsIndex]->headerEdits[$partName];
        }
        
        return $headerEdit;
    }
    
    public function getDocumentEdit($partName='word/document.xml') {
        if(is_null($this->words[$this->wordsIndex]->documentEdit)) {
            $index = $this->getPartsIndexByPartName(2,$partName);
            $document = $this->words[$this->wordsIndex]->parts[2][$index]['DOMElement'];
            if(isset($this->words[$this->wordsIndex]->blocks[2]) && isset($this->words[$this->wordsIndex]->blocks[2][$partName])) {
                $blocks = $this->words[$this->wordsIndex]->blocks[2][$partName];
            }else{
                $blocks = [];
            }
            
            foreach($this->words[$this->wordsIndex]->commentsEdit as $comments) {
                if($comments->partName === 'word/comments.xml') {
                    if(isset($comments->blocks)) {
                        $blocks = $this->my_array_merge($blocks,$comments->blocks);
                    }
                }
            }
            $documentEdit = new Document($this->words[$this->wordsIndex],$document,$blocks);
            $this->words[$this->wordsIndex]->documentEdit = $documentEdit;
            $this->words[$this->wordsIndex]->documentEdit->partName = $this->words[$this->wordsIndex]->parts[2][$index]['PartName'];
        }else{
            $documentEdit = $this->words[$this->wordsIndex]->documentEdit;
        }

        return $documentEdit;
    }
    
     /**
     * @var Document
     */
    public function getFooterEdit($partName='word/footer1.xml') {
        if(!isset($this->words[$this->wordsIndex]->footerEdits[$partName])) {
            $index = $this->getPartsIndexByPartName(23,$partName);
            $document = $this->words[$this->wordsIndex]->parts[23][$index]['DOMElement'];
            $blocks = $this->words[$this->wordsIndex]->blocks[23][$partName];//footer not add comment
            $blocks = $blocks?$blocks:[];
//             var_dump($partName,$document,$blocks);
            $footerEdit = new Footer($this->words[$this->wordsIndex],$document,$blocks);
            $this->words[$this->wordsIndex]->footerEdits[$partName] = $footerEdit;
            $this->words[$this->wordsIndex]->footerEdits[$partName]->partName = $this->words[$this->wordsIndex]->parts[23][$index]['PartName'];
        }else{
            $footerEdit = $this->words[$this->wordsIndex]->footerEdits[$partName];
        }
        return $footerEdit;
    }
    
    private function getPartsIndexByPartName($type,$partName) {
        static $data = [];
        
        if(!isset($data[$partName])) {
            $data = [];
            foreach($this->words[$this->wordsIndex]->parts as $part) {
                foreach($part as $index => $val) {
                    $data[$val['PartName']] = $index;
                }
            }
        }
        
        return $data[$partName];
    }
    
    /**
     * @return Styles
     */
    public function getStylesEdit() {
        $stylesEdit = $this->words[$this->wordsIndex]->stylesEdit;
        if(is_null($stylesEdit)) {
            $document = $this->words[$this->wordsIndex]->parts[4][0]['DOMElement'];
            $stylesEdit = new Styles($this->words[$this->wordsIndex],$document);
            $this->words[$this->wordsIndex]->stylesEdit = $stylesEdit;
            $this->words[$this->wordsIndex]->stylesEdit->partName = $this->words[$this->wordsIndex]->parts[4][0]['PartName'];
        }
        
        return $stylesEdit;
    }

    /**
     * @return Numbering
     */
    public function getNumberingEdit() {
        $numberingEdit = $this->words[$this->wordsIndex]->numberingEdit;
        if(is_null($numberingEdit) && isset($this->words[$this->wordsIndex]->parts[12])) {
            $document = $this->words[$this->wordsIndex]->parts[12][0]['DOMElement'];
            $numberingEdit = new Numbering($this->words[$this->wordsIndex],$document);
            $this->words[$this->wordsIndex]->numberingEdit = $numberingEdit;
            $this->words[$this->wordsIndex]->numberingEdit->partName = $this->words[$this->wordsIndex]->parts[12][0]['PartName'];
        }
        
        return $numberingEdit;
    }
    
    public function saveAs($fileName,$remainComments=false,$autoClean=true)
    {
        $tempFileName = $this->words[$this->wordsIndex]->save($remainComments);
        if (file_exists($fileName)) {
            unlink($fileName);
        }
        
        file_put_contents($fileName, file_get_contents($tempFileName));
        unlink($tempFileName);

        if($autoClean === true) {
            $this->free();
        }
    }
    
    public function saveAsContent() {
        $tempFileName = $this->words[$this->wordsIndex]->save();
        $content = file_get_contents($tempFileName);
        unlink($tempFileName);
        return $content;
    }
    
    public function saveAsToPathForTrace($dir,$baseName)
    {
        static $idx = 0;
        $word = $this->words[$this->wordsIndex];
        $tempFileName = $word->saveForTrace();
        
        $fileName = $dir.'/'.$baseName.'-'.str_pad($idx++,3,"0",STR_PAD_LEFT).'.docx';
        
        if (file_exists($fileName)) {
            unlink($fileName);
        }
        copy($tempFileName, $fileName);
        
        $WordProcessor = new WordProcessor();
        $WordProcessor->isForTrace = true;
        $WordProcessor->load($fileName);
        $WordProcessor->saveAs($fileName);
    }
    
    public function setChartValue($name,$fileName)
    {
        $reader = new Word();
        $reader->load($fileName);
        $this->words[++$this->wordsIndex] = $reader;
        $documentEdit = $this->getDocumentEdit();

        $documentChart = $documentEdit->getDocumentChart();

        $chartparts = $this->words[$this->wordsIndex]->getChartParts();
        $embeddings = $this->words[$this->wordsIndex]->getChartEmbeddings();

        $this->words[--$this->wordsIndex]->setChartParts($chartparts);
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setDocumentChart($name,$documentChart);
        $this->words[$this->wordsIndex]->updateChartRel();
        $this->words[$this->wordsIndex]->setContentTypes();
        $this->setEmbeddings();
   
    }

    public function setEmbeddings(){
        $this->words[$this->wordsIndex]->parts[13];
        foreach($this->words[$this->wordsIndex]->parts[13] as $part){
            if(!empty($part['embeddings'])){
                preg_match('/(\d+)/',$part['PartName'],$match);
                $fileName = preg_replace('/(\d+)/',$match[1],$part['embeddings']['name']);
                $this->words[$this->wordsIndex]->zip->addFromString('word/embeddings/'.$fileName, $part['embeddings']['xml']);
            }
        }
    }
    
    public function my_array_merge($arr,$arr2) {
        foreach($arr2 as $key => $val) {
            $arr[$key] = $val;
        }
        
        return $arr;
    }

    public function getMedies() {
        $word = $this->words[$this->wordsIndex];
        $numFiles = $word->zip->numFiles;
        $showList = [];
        for ($i = 0; $i < $numFiles; $i++) {
            $name = $word->zip->getNameIndex($i);
            if(strpos($name, 'media') > 0) {
                $content = $word->zip->getFromIndex($i);
                $showList['medias'][] = [
                    'md5' => md5($content),
                    'name' => $word->zip->getNameIndex($i),
                    'content' => $content,
                ];
            }
        }
        
        return $showList['medias'];
    }
    
    public function showMedies() {
        $showList = $this->getMedies();
        
        foreach($showList as $medias) {
            foreach($medias as $media) {
                var_dump($media);
                echo '<img src="data:image/jpeg;base64,'.base64_encode($media['content']).'"/><br/>';
            }
            
        }
    }
    
    public function getBlockList() {
        /**
         * @var Word $word
         */
        $word = $this->words[$this->wordsIndex];
        $tree = [];
        foreach($word->parts as $type => $parts) {
            
            switch ($type) {
                case 2:
                    $func = 'getDocumentEdit';
                    break;
                case 22:
                    $func = 'getHeaderEdit';
                    break;
                case 23:
                    $func = 'getFooterEdit';
                    break;
                default:
                    continue 2;
                    break;
            }
            
            foreach($parts as $part) {
                /**
                 * @var Document $document
                 */
                $document = $this->$func($part['PartName']);
                $tree[$type][] = $document->getBlockTree();
            }
        }
        
        return $tree;
    }

    public function getUsedCommentIds() {
        /**
         * @var Word $word
         */
        $word = $this->words[$this->wordsIndex];
        static $r = null;

        if($r !== null) {
            return $r;
        }

        $r = [];

        foreach($word->parts as $type => $parts) {
            
            switch ($type) {
                case 2:
                    $func = 'getDocumentEdit';
                    break;
                case 22:
                    $func = 'getHeaderEdit';
                    break;
                case 23:
                    $func = 'getFooterEdit';
                    break;
                default:
                    continue 2;
                    break;
            }
            
            foreach($parts as $part) {
                /**
                 * @var Document $document
                 */
                $document = $this->$func($part['PartName']);
                foreach($document->commentsblocks as $id => $comment) {
                    if(isset($document->usedBlock[$comment])) {
                        $r[$id] = 1;
                    }
                }
            }
        }
        
        return $r;
    }

    public function free(){
        foreach($this->words as $key => $word) {
            $word->free();
            unset($this->words[$key]);
        }
    }

    //删除tmp下文件
    public function deleteTmpFile(){
        $word = $this->words[$this->wordsIndex];
        $word->deleteTmpFile();
    }
}
