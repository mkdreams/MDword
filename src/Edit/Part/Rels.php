<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;
use MDword\Common\Build;

class Rels extends PartBase
{
    public $partInfo = null;
    
    protected $relationshipTypes =
    //--RELATIONSHIPTYPES--array (
  0 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  1 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  2 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  3 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  4 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  5 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  6 => 'http://schemas.microsoft.com/office/2011/relationships/commentsExtended',
  7 => 'http://schemas.microsoft.com/office/2011/relationships/people',
  8 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
  9 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  10 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings',
  11 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  12 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml',
  13 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  14 => 'http://schemas.microsoft.com/office/2016/09/relationships/commentsIds',
  15 => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
)//--RELATIONSHIPTYPES--
    ;
    
    public function __construct($word,\DOMDocument $DOMDocument) {
        parent::__construct($word);
        $this->DOMDocument = $DOMDocument;
        
        if(MDWORD_DEBUG) {
            $Relationships = $this->DOMDocument->getElementsByTagName('Relationship');
            foreach ($Relationships as $Relationship) {
                if(!in_array($type = $this->getAttr($Relationship, 'Type'), $this->relationshipTypes)) {
                    $this->relationshipTypes[] = $type;
                }
            }
            $build = new Build();
            $build->replace('RELATIONSHIPTYPES', $this->relationshipTypes, __FILE__);
        }
    }
    
    public function replace($rid,$file) {
        $Relationships = $this->DOMDocument->getElementsByTagName('Relationship');
        $length = $Relationships->length;
        foreach ($Relationships as $Relationship) {
            if($Relationship->getAttribute('Id') === $rid) {
                $type = $this->getAttr($Relationship, 'Type');
                switch ($type) {
                    case $this->relationshipTypes[2]:
                        $target = 'media/image'.++$length.'.png';
                        break;
                }
                
                //删除旧文件
                $oldValue = $this->partInfo['dirname'].'/'.$Relationship->getAttribute('Target');
                $this->word->zip->deleteName($oldValue);
                
                //替换
                $Relationship->setAttribute('Target',$target);
                $target = $this->partInfo['dirname'].'/'.$target;
                $this->word->zip->addFromString($target, file_get_contents($file));
            }
        }
    }
    
    public function getTarget($rid=null) {
        $Relationships = $this->DOMDocument->getElementsByTagName('Relationship');
        foreach ($Relationships as $Relationship) {
            if(is_null($rid)) {
                return $this->partInfo['dirname'].'/'.$Relationship->getAttribute('Target');
            }
            
            if($Relationship->getAttribute('Id') === $rid) {
                return $this->partInfo['dirname'].'/'.$Relationship->getAttribute('Target');
            }
        }
        
        return null;
    }
}
