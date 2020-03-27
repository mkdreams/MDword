<?php
namespace PhpOffice\PhpWord\Processor\Common;

class PartBase
{
    protected $DOMDocument;
    
    /**
     * 
     * @var \PhpOffice\PhpWord\Processor\Common\View
     */
    protected $view;
    
    protected $rootPath;
    
    public function __construct() {
        $this->view = new View();
        $this->rootPath = dirname(__DIR__);
    }
    
    public function parse() {
        
    }
}
