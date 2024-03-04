<?php
namespace MDword\Read\Part;

use MDword\Common\PartBase;
use MDword\Common\Build;

class ContentTypes extends PartBase
{
    protected $defaults = [];
    protected $overrides = [];
    
    protected $contentTypes =
    //--CONTENTTYPES--
array (
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
  11 => 'application/octet-stream',
  12 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
  13 => 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  14 => 'image/png',
  15 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
  16 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml',
  17 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml',
  18 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml',
  19 => 'rels',
  20 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml',
  21 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml',
  22 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  23 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  24 => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  25 => 'application/vnd.ms-office.chartstyle+xml',
  26 => 'application/vnd.ms-office.chartcolorstyle+xml',
  27 => 'image/jpeg',
  28 => 'image/x-wmf',
  29 => 'application/vnd.openxmlformats-officedocument.oleObject',
  30 => 'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml',
  31 => 'image/x-ms-bmp',
)//--CONTENTTYPES--
    ;
    /**
     * @param \DOMDocument $DOMDocument
     */
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
        
        if(MDWORD_DEBUG) {
            $build = new Build();
            $build->replace('CONTENTTYPES', $this->contentTypes, __FILE__);
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
                $Extension = $item->getAttribute('Extension');
                $this->defaults[$Extension] = ['Extension'=>$Extension,'ContentType'=>$pos];
                break;
            case 'Override':
                $PartName = ltrim($item->getAttribute('PartName'),'/');
                $this->overrides[$PartName] = ['PartName'=>$PartName,'ContentType'=>$pos];
                break;
        }
    }
    
    public function addDefault($Extension,$ContentType) {
        if(!isset($this->defaults[$Extension])) {
            $pos = array_search($ContentType,$this->contentTypes);
            if($pos === false) {
                $this->word->log->writeLog('content type not find! type:'.$ContentType);
            }else{
                $this->defaults[$Extension] = ['Extension'=>$Extension,'ContentType'=>$pos];
            }
            
            $node = $this->createNodeByXml('<Default Extension="'.$Extension.'" ContentType="'.$ContentType.'"/>');
            $Types = $this->DOMDocument->getElementsByTagName('Types')->item(0);
            $this->insertBefore($node, $Types->firstChild);
        }
    }
    
    public function addOverride($PartName='word/document.xml',$ContentTypeIdx=2) {
        if(is_string($ContentTypeIdx)) {
            $pos = array_search($ContentTypeIdx,$this->contentTypes);
            if($pos === false) {
                $this->contentTypes[] = $ContentTypeIdx;
                $pos = sizeof($this->contentTypes) - 1;
            }
            $ContentTypeIdx = $pos;
        }
        $this->overrides[$PartName] = ['PartName'=>$PartName,'ContentType'=>$ContentTypeIdx];

        $overrides = $this->DOMDocument->getElementsByTagName('Override');
        $lastOverride = $overrides->item($overrides->length-1);
        $copy = clone $lastOverride;
        $this->setAttr($copy,'PartName',$PartName,'');
        $this->setAttr($copy,'ContentType',$this->contentTypes[$ContentTypeIdx],'');
        $this->insertAfter($copy,$lastOverride);
    }

    public function setContent_types($newOverrides){
        $this->overrides = array_merge($this->overrides,$newOverrides);
        $xlsDefault = ['Extension' => 'xlsx','ContentType' =>11];
        foreach($this->defaults as $default){
            $defaultArr[] = $default['Extension'];
        }
        if(!array_search('xlsx',$defaultArr)){
            $this->defaults[] = $xlsDefault;
            $defaults = $this->DOMDocument->getElementsByTagName('Default');
            $copy = clone $defaults[0];
            $copy->setAttribute('Extension',$xlsDefault['Extension']);
            $copy->setAttribute('ContentType',$this->contentTypes[$xlsDefault['ContentType']]);
            $defaults[0]->parentNode->appendChild($copy);
        }
    }
}
