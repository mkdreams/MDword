<?php
namespace MDword;

use MDword\Read\Word2007;
use MDword\Edit\Part\Document;

class WordProcessor
{
    private $wordsIndex = -1;
    private $words = [];
    
    public function load($zip) {
        $reader = new Word2007();
        $reader->load($zip);
        $this->words[++$this->wordsIndex] = $reader;
        return $this->words[$this->wordsIndex];
    }
    
    public function setValue($name, $value) {
        $document = $this->words[$this->wordsIndex]->parts[2][0]['DOMElement'];
        $documentEdit = new Document($document);
        $documentEdit->setValue($name, $value);
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
