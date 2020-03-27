<?php
namespace MDword;

use MDword\Read\Word2007;

class WordProcessor
{
    private $words = null;
    
    public function load($zip) {
        $reader = new Word2007();
        $reader->load($zip);
        $this->words[] = $reader;
    }
}
