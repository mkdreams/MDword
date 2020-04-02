<?php
namespace MDword;

use MDword\Edit\Part\Document;
use MDword\Read\Word;
use MDword\Edit\Part\Comments;

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
        
        $comments = $this->words[$this->wordsIndex]->parts[15][0]['DOMElement'];
        $this->words[$this->wordsIndex]->commentsEdit = new Comments($this->words[$this->wordsIndex],$comments);
        $this->words[$this->wordsIndex]->commentsEdit->partName = $this->words[$this->wordsIndex]->parts[15][0]['PartName'];
        $this->words[$this->wordsIndex]->commentsEdit->word = $this->words[$this->wordsIndex];
        
        return $this->words[$this->wordsIndex];
    }
    
    public function setValue($name, $value) {
        $documentEdit = $this->words[$this->wordsIndex]->documentEdit;
        if(is_null($documentEdit)) {
            $document = $this->words[$this->wordsIndex]->parts[2][0]['DOMElement'];
            $documentEdit = new Document($this->words[$this->wordsIndex],$document,$this->words[$this->wordsIndex]->commentsEdit->blocks);
            $this->words[$this->wordsIndex]->documentEdit = $documentEdit;
            $this->words[$this->wordsIndex]->documentEdit->partName = $this->words[$this->wordsIndex]->parts[2][0]['PartName'];
        }
        $documentEdit->setValue($name, $value);
    }
    
    public function setImageValue($name, $value) {
        $documentEdit = $this->words[$this->wordsIndex]->documentEdit;
        if(is_null($documentEdit)) {
            $document = $this->words[$this->wordsIndex]->parts[2][0]['DOMElement'];
            $documentEdit = new Document($this->words[$this->wordsIndex],$document,$this->words[$this->wordsIndex]->commentsEdit->blocks);
            $this->words[$this->wordsIndex]->documentEdit = $documentEdit;
            $this->words[$this->wordsIndex]->documentEdit->partName = $this->words[$this->wordsIndex]->parts[2][0]['PartName'];
        }
        $documentEdit->setValue($name, $value,'image');
    }
    
    public function clone($name,$count=1) {
        $documentEdit = $this->words[$this->wordsIndex]->documentEdit;
        if(is_null($documentEdit)) {
            $document = $this->words[$this->wordsIndex]->parts[2][0]['DOMElement'];
            $documentEdit = new Document($this->words[$this->wordsIndex],$document,$this->words[$this->wordsIndex]->commentsEdit->blocks);
            $this->words[$this->wordsIndex]->documentEdit = $documentEdit;
            $this->words[$this->wordsIndex]->documentEdit->partName = $this->words[$this->wordsIndex]->parts[2][0]['PartName'];
        }
        $documentEdit->clone($name, $count);
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
    
    
}
