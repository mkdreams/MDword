<?php
namespace MDword\XmlTemple;

use PhpOffice\PhpWord\Media;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\Writer\AbstractWriter;
use PhpOffice\PhpWord\Writer\WriterInterface;
use MDword\Edit\Part\Document;

class XmlFromPhpword extends AbstractWriter implements WriterInterface
{
    private $document = null;
    /**
     * @param \PhpOffice\PhpWord\PhpWord
     * @param \MDword\Edit\Part\Document
     */
    public function __construct(PhpWord $phpWord = null,Document $document = null)
    {
        // Assign PhpWord
        $this->phpWord = $phpWord;
        $this->document = $document;

        // Create parts
        $this->parts = array(
            'ContentTypes'   => '[Content_Types].xml',
            'Rels'           => '_rels/.rels',
            'DocPropsApp'    => 'docProps/app.xml',
            'DocPropsCore'   => 'docProps/core.xml',
            'DocPropsCustom' => 'docProps/custom.xml',
            'RelsDocument'   => 'word/_rels/document.xml.rels',
            'Document'       => 'word/document.xml',
            'Comments'       => 'word/comments.xml',
            'Styles'         => 'word/styles.xml',
            'Numbering'      => 'word/numbering.xml',
            'Settings'       => 'word/settings.xml',
            'WebSettings'    => 'word/webSettings.xml',
            'FontTable'      => 'word/fontTable.xml',
            'Theme'          => 'word/theme/theme1.xml',
            'RelsPart'       => '',
            'Header'         => '',
            'Footer'         => '',
            'Footnotes'      => '',
            'Endnotes'       => '',
            'Chart'          => '',
        );
        foreach (array_keys($this->parts) as $partName) {
            $partClass = 'PhpOffice\\PhpWord\\Writer\\Word2007\\Part\\' . $partName;
            if (class_exists($partClass)) {
                /** @var \PhpOffice\PhpWord\Writer\Word2007\Part\AbstractPart $part Type hint */
                $part = new $partClass();
                $part->setParentWriter($this);
                $this->writerParts[strtolower($partName)] = $part;
            }
        }
    }
    
    private function writeStyles() {
        $skipTags = ['Normal'=>1,'Footnote Reference'=>1];
        $stylesEdit = $this->document->word->wordProcessor->getStylesEdit();
        
        $stylesXml = $this->getPartXml('styles');
        $stylesNode = $stylesEdit->createNodeByXml($stylesXml,function($documentElement) {
            return $documentElement;
        });
        
        $names = [];
        $styles = $stylesNode->getElementsByTagName('style');
        foreach($styles as $key => $style) {
            $nameNode = $style->getElementsByTagName('name')->item(0);
            $name = $stylesEdit->getAttr($nameNode, 'val');
            if(isset($skipTags[$name])) {
                continue;
            }
            
            $newName = 'PHPWORD'.$key.$name;
            $names[$name] = $newName;
            $stylesEdit->setAttr($nameNode, 'val', $newName);
            
            $copy = clone $style;
            $stylesEdit->DOMDocument->documentElement->appendChild($copy);
        }
        
        return $names;
    }
    
    private function writeImages() {
        $sectionMedia = Media::getElements('section');
        $rids = [];
        foreach($sectionMedia as $media) {
            if($media['type'] === 'image') {
                $refInfo = $this->document->updateRef($media['source'],null,MDWORD_IMG);
                $rids['r:id="rId'.($media['rID']+6).'"'] = 'r:id="'.$refInfo['rId'].'"';
            }
        }
        
        return $rids;
    }
    
    private function getPartXml($partName) {
        $xml = $this->writerParts[$partName]->write();
        return $xml;
    }
    
    public function createNodesByBodyXml() {
        $styleNames = $this->writeStyles();
        $imageRids = $this->writeImages();
        
        $xml = $this->getPartXml('document');
        $xml = str_replace(
            array_merge(array_keys($styleNames),array_keys($imageRids)),
            array_merge(array_values($styleNames),array_values($imageRids)),
            $xml);
        $body = $this->document->createNodeByXml($xml,function($documentElement) {
            return $documentElement->getElementsByTagName('body')->item(0);
        });
        
        return $body->childNodes;
    }

    /**
     * Save document by name.
     *
     * @param string $filename
     */
    public function save($filename = null)
    {
        $filename = $this->getTempFile($filename);
        $zip = $this->getZipArchive($filename);
        $phpWord = $this->getPhpWord();
//         var_dump($this->parts);exit;

        // Content types
        $this->contentTypes['default'] = array(
            'rels' => 'application/vnd.openxmlformats-package.relationships+xml',
            'xml'  => 'application/xml',
        );

        // Add section media files
        $sectionMedia = Media::getElements('section');
        if (!empty($sectionMedia)) {
            $this->addFilesToPackage($zip, $sectionMedia);
            $this->registerContentTypes($sectionMedia);
            foreach ($sectionMedia as $element) {
                $this->relationships[] = $element;
            }
        }

        // Add header/footer media files & relations
        $this->addHeaderFooterMedia($zip, 'header');
        $this->addHeaderFooterMedia($zip, 'footer');

        // Add header/footer contents
        $rId = Media::countElements('section') + 6; // @see Rels::writeDocRels for 6 first elements
        $sections = $phpWord->getSections();
//         var_dump($this->parts);exit;
        foreach ($sections as $section) {
            $this->addHeaderFooterContent($section, $zip, 'header', $rId);
            $this->addHeaderFooterContent($section, $zip, 'footer', $rId);
        }

        $this->addNotes($zip, $rId, 'footnote');
        $this->addNotes($zip, $rId, 'endnote');
        $this->addComments($zip, $rId);
        $this->addChart($zip, $rId);

//         var_dump($this->parts);exit;
//         echo $this->writerParts['document']->write();exit;
        echo $this->writerParts['styles']->write();exit;
        
        // Write parts
        foreach ($this->parts as $partName => $fileName) {
            if ($fileName != '') {
                $zip->addFromString($fileName, $this->getWriterPart($partName)->write());
            }
        }

        // Close zip archive and cleanup temp file
        $zip->close();
        $this->cleanupTempFile();
    }
}
