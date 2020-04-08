<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Charts extends PartBase
{
    public $blocks = [];
    public function __construct($word,\DOMDocument $DOMDocument) {
        parent::__construct($word);
    }
}
