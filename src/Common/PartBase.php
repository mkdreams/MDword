<?php
namespace MDword\Common;

class PartBase
{
    protected $DOMDocument;
    
    /**
     * 
     * @var \MDword\Common\View
     */
    protected $view;
    
    protected $rootPath;
    
    public function __construct() {
        $this->view = new View();
        $this->rootPath = dirname(__DIR__);
    }
    
    public function parse() {
        
    }
    
    public function __get($name) {
        return $this->$name;
    }
}
