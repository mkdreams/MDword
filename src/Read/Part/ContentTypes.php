<?php
namespace MDword\Read\Part;

use MDword\Common\PartBase;
use MDword\Common\View;

class ContentTypes extends PartBase
{
    protected $defaults = [];
    protected $overrides = [];
    protected $partNames = [];
    
    protected $contentTypes = array (
        0 => 'application/vnd.openxmlformats-package.relationships+xml',
        1 => 'application/xml',
        2 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
        3 => 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml',
        4 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
        5 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',
        6 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml',
        7 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml',
        8 => 'application/vnd.openxmlformats-officedocument.theme+xml',
        9 => 'application/vnd.openxmlformats-package.core-properties+xml',
        10 => 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    );
    
    public function __construct(\DOMDocument $DOMDocument) {
        parent::__construct();
        
        $this->DOMDocument = $DOMDocument;
        $this->parse();
    }
    
    public function parse() {
        $Types = $this->DOMDocument->getElementsByTagName('Types')->item(0);
        $childrens = $Types->getElementsByTagName('*');
        foreach($childrens as $children) {
            $this->paseItem($children);
        }
    }
    
    public function getPartNames() {
        return $this->partNames;
    }
    
    private function paseItem(\DOMElement $item) {
        $ContentType = $item->getAttribute('ContentType');
        $pos = array_search($ContentType,$this->contentTypes);
        if($pos === false) {
            $this->contentTypes[] = $ContentType;
            $pos = sizeof($this->contentTypes) - 1;
        }
        
        switch ($item->tagName) {
            case 'Default':
                $this->defaults[] = ['Extension'=>$item->getAttribute('Extension'),'ContentType'=>$pos];
                break;
            case 'Override':
                $PartName = ltrim($item->getAttribute('PartName'),'/');
                $this->overrides[] = ['PartName'=>$PartName,'ContentType'=>$pos];
                $this->partNames[] = $PartName;
                break;
        }
    }
    
    public function display()
    {
        $this->view->assign('defaults',$this->defaults);
        $this->view->assign('overrides',$this->overrides);
        return $this->view->fetch('ContentTypes');
    }
}
