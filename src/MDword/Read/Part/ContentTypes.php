<?php
namespace PhpOffice\PhpWord\Processor\Read\Part;

use PhpOffice\PhpWord\Processor\Common\PartBase;
use PhpOffice\PhpWord\Processor\Common\View;

class ContentTypes extends PartBase
{
    private $defaults = [];
    private $overrides = [];
    
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
        
        
//         echo $this->display();
//         exit;
//         var_dump($this->defaults,$this->overrides);exit;
    }
    
    private function paseItem(\DOMElement $item) {
        switch ($item->tagName) {
            case 'Default':
                $this->defaults[] = ['Extension'=>$item->getAttribute('Extension'),'ContentType'=>$item->getAttribute('ContentType')];
                break;
            case 'Override':
                $this->overrides[] = ['PartName'=>$item->getAttribute('PartName'),'ContentType'=>$item->getAttribute('ContentType')];
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
