<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;
use MDword\XmlTemple\XmlFromPhpword;

class Styles extends PartBase
{
    private $styles = null;
    public function __construct($word,\DOMDocument $DOMDocument) {
        parent::__construct($word);
        
        $this->DOMDocument = $DOMDocument;
        $this->initNameSpaces();
    }

    public function getStyleById($id) {
        if(is_null($this->styles)) {
            $tempStyles = $this->DOMDocument->getElementsByTagName('style');
            $this->styles = [];
            foreach($tempStyles as $style) {
                $styleId = $this->getAttr($style, 'styleId');
                $this->styles[$styleId] = $style;
            }
        }

        return $this->styles[$id];
    }
}
