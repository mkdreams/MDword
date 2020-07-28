<?php
namespace MDword;

use MDword\Edit\Part\Document;
use MDword\Read\Word;
use MDword\Edit\Part\Comments;
use MDword\Common\Bind;

class WordProcessor
{
    private $wordsIndex = -1;
    private $words = [];
    
    public function __construct() {
        require_once(dirname(__FILE__).DIRECTORY_SEPARATOR.'config'.DIRECTORY_SEPARATOR.'main.php');
    }
    
    public function load($zip) {
        $reader = new Word();
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
//         $comments = $this->words[$this->wordsIndex]->parts[15][0]['DOMElement'];
//         $this->words[$this->wordsIndex]->commentsEdit = new Comments($this->words[$this->wordsIndex],$comments);
//         $this->words[$this->wordsIndex]->commentsEdit->partName = $this->words[$this->wordsIndex]->parts[15][0]['PartName'];
//         $this->words[$this->wordsIndex]->commentsEdit->word = $this->words[$this->wordsIndex];
        
        $this->getDocumentEdit();
        
        return $this->words[$this->wordsIndex];
    }
    
    /**
     * @return \MDword\Common\Bind
     */
    public function getBind($data) {
        $bind = new Bind($this,$data);
        return $bind;
    }
    
    public function setValue($name, $value) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($name, $value);
    }
    
    public function setValues($values,$pre='') {
        foreach ($values as $index => $valueArr) {
            foreach($valueArr as $name => $value) {
                if(is_array($value)) {
                    $this->setValues($value,'#'.$index);
                }else{
                    $this->setValue($name.$pre.'#'.$index, $value);
                }
            }
        }
    }
    
    /**
     * delete p include the block 
     * @param string $name
     */
    public function deleteP(string $name) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($name, 'p','delete');
    }
    
    /**
     * delete tr include the block
     * @param string $name
     */
    public function deleteTr(string $name) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($name, 'tr','delete');
    }
    
    /**
     * delete block
     * @param string $name
     */
    public function delete(string $name) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($name, '','text');
    }
    
    public function setImageValue($name, $value) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->update([],$name,$value,'image');
    }
    
    /**
     * @param string $name
     * ['text','link']
     * @param array $value
     */
    public function setLinkValue($name, $value) {
        $documentEdit = $this->getDocumentEdit();
        
        $documentEdit->setValue($name, $value[0],'text');
        $documentEdit->setValue($name, $value[1],'link');
    }
    
    /**
     * @param string $name
     * @param array $datas
     * change value ['A1',9,'set']
     * extention range ['$A$1:$A$5','$A$1:$A$10','ext']
     */
    public function setExcelValues($name='',$datas=[]) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($name, $datas, 'excel');
    }
    
    /**
     * clone p
     * @param string $name
     * @param int $count
     */
    public function cloneP($name,$count=1) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($name, $count, 'cloneP');
    }
    /**
     * clone
     * @param string $name
     * @param int $count
     */
    public function clone($name,$count=1) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($name, $count, 'clone');
    }
    /**
     * clone
     * @param string $name
     * @param int $count
     */
    public function cloneTo($nameTo,$name) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($nameTo, $name, 'cloneTo');
    }
    
    
    public function setBreakValue($name, $value) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($name, $value,'break');
    }
    
    
    public function setBreakPageValue($name, $value=1) {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->setValue($name, $value,'breakpage');
    }
    
    /**
     * update toc
     */
    public function updateToc() {
        $documentEdit = $this->getDocumentEdit();
        $documentEdit->updateToc();
    }
    
    
    private function getDocumentEdit() {
        $documentEdit = $this->words[$this->wordsIndex]->documentEdit;
        if(is_null($documentEdit)) {
            $document = $this->words[$this->wordsIndex]->parts[2][0]['DOMElement'];
            $blocks = [];
            foreach($this->words[$this->wordsIndex]->commentsEdit as $coments) {
                if($coments->partName === 'word/comments.xml') {
                    $blocks = $coments->blocks;
                }
            }
            $documentEdit = new Document($this->words[$this->wordsIndex],$document,$blocks);
            $this->words[$this->wordsIndex]->documentEdit = $documentEdit;
            $this->words[$this->wordsIndex]->documentEdit->partName = $this->words[$this->wordsIndex]->parts[2][0]['PartName'];
        }
        return $documentEdit;
    }
    
    public function saveAs($fileName)
    {
        $tempFileName = $this->words[$this->wordsIndex]->save();
        
        if (file_exists($fileName)) {
            unlink($fileName);
        }
        
        copy($tempFileName, $fileName);
        unlink($tempFileName);
    }
    
    public function setChartValue($name='',$fileName)
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
    
    public function showMedies() {
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
        
        foreach($showList as $medias) {
            foreach($medias as $media) {
                var_dump($media);
                echo '<img src="data:image/jpeg;base64,'.base64_encode($media['content']).'"/><br/>';
            }
            
        }
    }
}
