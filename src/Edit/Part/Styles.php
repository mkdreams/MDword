<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;
use MDword\XmlTemple\XmlFromPhpword;

class Styles extends PartBase
{
    public function __construct($word,\DOMDocument $DOMDocument) {
        parent::__construct($word);
        
        $this->DOMDocument = $DOMDocument;
    }
}
