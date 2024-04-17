<?php
namespace MDword\Read;

use MDword\Edit\Part\Document;
use MDword\Edit\Part\Comments;
use MDword\Read\Part\ContentTypes;
use MDword\Common\Log;
use MDword\WordProcessor;
use MDword\Edit\Part\Styles;
use MDword\Edit\Part\Footer;
use MDword\Edit\Part\Header;

class Word 
{
    public $zip = null;
    /**
     * @var Log
     */
    public $log = null;
    /**
     * @var ContentTypes
     */
    public $Content_Types = null;
    
    private $tempDocumentFilename = null;
    /**
     * @var Header
     */
    public $headerEdits = [];
    /**
     * @var Document
     */
    public $documentEdit = null;
    /**
     * @var Footer
     */
    public $footerEdits = [];
    /**
     * @var Comments
     */
    public $commentsEdit = null;
    /**
     * @var Styles
     */
    public $stylesEdit = null;
    
    public $parts = [];
    
    public $files = [];
    
    public $blocks = [];
    
    /**
     * @var array
     */
    public $needUpdateParts = []; 
    
    /**
     * @var WordProcessor
     */
    public $wordProcessor = null;
    
    
    public function __construct() {
        $this->log = new Log();
    }
    
    
    public function load($archive) {
        if(!is_file($archive)) {
            throw new \Exception('Template not exist!');
        }

        $this->tempDocumentFilename = tempnam($this->getTempDir(), 'MDword');
        if (false === $this->tempDocumentFilename) {
            throw new \Exception('temp path make failed!');
        }
        
        
        if (false === file_put_contents($this->tempDocumentFilename, file_get_contents($archive)) || !is_file($this->tempDocumentFilename)) {
            throw new \Exception($archive.' copy file failed!');
        }
        
        $this->zip = new \ZipArchive();
        $this->zip->open($this->tempDocumentFilename);
        
        $this->read();
        
        //update needUpdateParts
        if(isset($this->parts[22]) && isset($this->blocks[22])) {
            foreach($this->blocks[22] as $partName => $value) {
                $this->needUpdateParts[$partName] = ['func'=>'getHeaderEdit','partName'=>$partName];
            }
        }
        
        $this->needUpdateParts['word/document.xml'] = ['func'=>'getDocumentEdit','partName'=>'word/document.xml'];
        
        if(isset($this->parts[23])  && isset($this->blocks[23])) {
            foreach($this->blocks[23] as $partName => $value) {
                $this->needUpdateParts[$partName] = ['func'=>'getFooterEdit','partName'=>$partName];
            }
        }
    }
    
    private function read() {
        $this->Content_Types = new ContentTypes($this->getXmlDom('[Content_Types].xml'));
        $this->Content_Types->word = $this;
        foreach ($this->Content_Types->overrides as $part) {
            if(in_array($part['ContentType'], [2,22,23])) {//document footer header
                $standardXmlFunc = function($xml) use($part) {
                    $xml = $this->standardXml($xml,$part['ContentType'],$part['PartName']);
                    return $xml;
                };
            }else{
                $standardXmlFunc = null;
            }
            
            if($part['ContentType'] === 14) {//image/png
                $this->parts[$part['ContentType']][] = ['PartName'=>$part['PartName'],'DOMElement'=>$part['PartName']];
            }else{
                $this->parts[$part['ContentType']][] = ['PartName'=>$part['PartName'],'DOMElement'=>$this->getXmlDom($part['PartName'],$standardXmlFunc)];
            }
        }
    }
    
    public static function getTempDir()
    {
        $tempDir = sys_get_temp_dir();
        
        if (!empty(self::$tempDir)) {
            $tempDir = self::$tempDir;
        }
        
        return $tempDir;
    }
    
    public function save($remainComments = false)
    {
        //clean before update
        $this->deleteComments($remainComments);
        
        //update Toc
        $this->documentEdit->updateToc();
        
        $this->deleteComments($remainComments);
        
        foreach($this->parts as $type => $list ) {
            foreach($list as $part) {
                if(is_object($part['DOMElement'])) {
                    if(empty($part['DOMElement']->documentElement)){
                        continue;
                    }
                    //delete space before break page
                    if($type === 2) {
                        $this->zip->addFromString($part['PartName'], $this->SimSunExtBSupport($this->autoDeleteSpaceBeforeBreakPage($part['DOMElement']->saveXML())));
                    }else{
                        //add MDword flag
                        if($type === 9) {
                            $description = $part['DOMElement']->getElementsByTagName('description')->item(0);
                            if(is_null($description)) {
                                $el = $part['DOMElement']->createElement('dc:description','Made with MDword');
                                if($coreProperties = $part['DOMElement']->getElementsByTagName('coreProperties')->item(0)) {
                                    $coreProperties->appendChild($el); 
                                }
                            }else{
                                $description->nodeValue='Made with MDword';
                            }

                        }
                        $this->zip->addFromString($part['PartName'], $this->SimSunExtBSupport($part['DOMElement']->saveXML()));
                    }
                }
            }
        }
        
        foreach($this->files as $part) {
            if(isset($part['this'])) {
                $this->zip->addFromString($part['PartName'], $part['this']->getContent());
            }else{
                $this->zip->deleteName($part['PartName']);
            }
        }
        
        $this->zip->addFromString('[Content_Types].xml', $this->Content_Types->DOMDocument->saveXML());
        
        if (false === $this->zip->close()) {
            throw new \Exception('Could not close zip file.');
        }
        
//         trigger_error('debug',E_USER_ERROR);exit;
        
        if(MDWORD_DEBUG === true) {
            $this->zip->open($this->tempDocumentFilename);
            $this->zip->extractTo(MDWORD_GENERATED_DIRECTORY);
            $this->zip->close();
        }
        
        return $this->tempDocumentFilename;
    }
    
    public function saveForTrace()
    {
        foreach($this->parts as $type => $list ) {
            foreach($list as $part) {
                if(is_object($part['DOMElement'])) {
                    //delete document end space
                    if($type === 2) {
                        $this->zip->addFromString($part['PartName'], $this->autoDeleteSpaceBeforeBreakPage($part['DOMElement']->saveXML()));
                    }else{
                        $this->zip->addFromString($part['PartName'], $part['DOMElement']->saveXML());
                    }
                }
            }
        }
        
        foreach($this->files as $part) {
            if(isset($part['this'])) {
                $this->zip->addFromString($part['PartName'], $part['this']->getContent());
            }else{
                $this->zip->deleteName($part['PartName']);
            }
        }
        
        $this->zip->addFromString('[Content_Types].xml', $this->Content_Types->DOMDocument->saveXML());
        
        
        $this->zip->open($this->tempDocumentFilename);
        $this->zip->extractTo(MDWORD_GENERATED_DIRECTORY);
        
        return $this->tempDocumentFilename;
    }
    
    
    public function standardXml($xml,$ContentType,$PartName) {
        $xml = preg_replace_callback('/\$[^$]*?\{[\s\S]+?\}/i', function($match){
            return preg_replace('/\s/', '', strip_tags($match[0]));
        }, $xml);
        
        static $commentId = 0;
        $nameToCommendId = [];
        $xml = preg_replace_callback('/(<[w|m]\:r[> ](?:(?!<[w|m]:r[> ])[\S\s])*?<[w|m]\:t[ ]{0,1}[^>]*?>)([^><]*?)(<\/[w|m]\:t>[\s\S]*?<\/[w|m]\:r>)/i', function($matchs) use(&$commentId,&$nameToCommendId,$ContentType,$PartName){
            return preg_replace_callback('/\$\{([\s\S]+?)\}/i', function($match) use(&$commentId,&$nameToCommendId,$matchs,$ContentType,$PartName){
                $name = $match[1];
                $length = strlen($name);
                if($name[$length-1] === '/') {
                    $name = trim($name,'/');
                    $this->blocks[$ContentType][$PartName]['r'.$commentId] = $name;
                    return $matchs[3].'<w:commentRangeStart w:id="r'.$commentId.'"/>'.$matchs[1].$match[0].$matchs[3].'<w:commentRangeEnd w:id="r'.$commentId++.'"/>'.$matchs[1];
                //end
                }elseif($name[0] === '/') {
                    $name = trim($name,'/');
                    return $match[0].$matchs[3].'<w:commentRangeEnd w:id="r'.$nameToCommendId[$name].'"/>'.$matchs[1];
                //start
                }else{
                    $name = trim($name,'/');
                    $this->blocks[$ContentType][$PartName]['r'.$commentId] = $name;
                    $nameToCommendId[$name] = $commentId;
                    return $matchs[3].'<w:commentRangeStart w:id="r'.$commentId++.'"/>'.$matchs[1].$match[0];
                }
                
            }, $matchs[0]);
        }, $xml);
        
        return $xml;
    }
    
    private function autoDeleteSpaceBeforeBreakPage($xml) {
        ini_set("pcre.backtrack_limit",-1);//回溯bug fixed,-1代表不做限制
        $xml = preg_replace_callback('/\<\/w\:p\>(?:\<w\:p\>(?:(?!\<\/w\:t\>|\<v\:imagedata).)+?\<\/w\:p\>|\<w\:p\/\>)*?\<w\:p[^>]*?\>\<w\:r\>\<w\:br w\:type\=\"page\"\/\>\<\/w\:r\>\<\/w\:p\>/i', function($match) {
            
            preg_match_all('/\<w\:p[^>]*?\>(.+?)\<\/w\:p\>/i', $match[0],$subMatch);
            if(!empty($subMatch[1])) {
                return implode('', $subMatch[1]).'</w:p>';
            }
            
            return '<w:r><w:br w:type="page"/></w:r></w:p>';
        }, $xml);
            
        $xml = preg_replace('/\<w\:p\/\>\<w\:sectPr\>/','<w:sectPr>',$xml);
        
        return $xml;
    }

    private function SimSunExtBSupport($xml) {
        //4字节支持字体转化
        preg_match_all('/\<w:rFonts[\s\S]*?\/\>/i',$xml,$matches);
        $font = empty($matches[0])?'<w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体"/></w:rPr>':'<w:rPr>'.end($matches[0]).'</w:rPr>';


        return preg_replace([
            '/\{%\{0\}%\}([\x{20000}-\x{2a6d6}|\x{2a700}-\x{2b734}|\x{2b740}-\x{2b81d}|\x{2b8b8}-\x{2b8b9}|\x{2bac7}-\x{2bac8}|\x{2bb5f}-\x{2bb60}|\x{2bb62}-\x{2bb63}|\x{2bb7c}-\x{2bb7d}|\x{2bb83}-\x{2bb84}|\x{2bc1b}-\x{2bc1c}|\x{2bd77}|\x{2bd87}|\x{2bdf7}|\x{2be29}|\x{2c029}-\x{2c02a}|\x{2c0a9}|\x{2c0ca}|\x{2c1d5}|\x{2c1d9}|\x{2c1f9}|\x{2c27c}|\x{2c288}|\x{2c2a4}|\x{2c317}|\x{2c35b}|\x{2c361}|\x{2c364}|\x{2c488}|\x{2c494}|\x{2c497}|\x{2c542}|\x{2c613}|\x{2c618}|\x{2c621}|\x{2c629}|\x{2c62b}-\x{2c62d}|\x{2c62f}|\x{2c642}|\x{2c64a}-\x{2c64b}|\x{2c72c}|\x{2c72f}|\x{2c79f}|\x{2c7c1}|\x{2c7fd}|\x{2c8d9}|\x{2c8de}|\x{2c8e1}|\x{2c8f3}|\x{2c907}|\x{2c90a}|\x{2c91d}|\x{2ca02}|\x{2ca0e}|\x{2ca7d}|\x{2caa9}|\x{2cb29}|\x{2cb2d}-\x{2cb2e}|\x{2cb31}|\x{2cb38}-\x{2cb39}|\x{2cb3b}|\x{2cb3f}|\x{2cb41}|\x{2cb4a}|\x{2cb4e}|\x{2cb5a}-\x{2cb5b}|\x{2cb64}|\x{2cb69}|\x{2cb6c}|\x{2cb6f}|\x{2cb73}|\x{2cb76}|\x{2cb78}|\x{2cb7c}|\x{2cbb1}|\x{2cbbf}-\x{2cbc0}|\x{2cbce}|\x{2cc56}|\x{2cc5f}|\x{2ccf5}-\x{2ccf6}|\x{2ccfd}|\x{2ccff}|\x{2cd02}-\x{2cd03}|\x{2cd0a}|\x{2cd8b}|\x{2cd8d}|\x{2cd8f}-\x{2cd90}|\x{2cd9f}-\x{2cda0}|\x{2cda8}|\x{2cdad}-\x{2cdae}|\x{2cdd5}|\x{2ce18}|\x{2ce1a}|\x{2ce23}|\x{2ce26}|\x{2ce2a}|\x{2ce7c}|\x{2ce88}|\x{2ce93}]+)/u',
            '/\{%\{1\}%\}([\x{0}|\x{d}|\x{100}|\x{102}-\x{112}|\x{114}-\x{11a}|\x{11c}-\x{12a}|\x{12c}-\x{143}|\x{145}-\x{147}|\x{149}-\x{14c}|\x{14e}-\x{151}|\x{154}-\x{15f}|\x{162}-\x{16a}|\x{16c}-\x{177}|\x{179}-\x{191}|\x{193}-\x{1cd}|\x{1cf}|\x{1d1}|\x{1d3}|\x{1d5}|\x{1d7}|\x{1d9}|\x{1db}|\x{1dd}-\x{1f8}|\x{1fa}-\x{250}|\x{252}-\x{260}|\x{262}-\x{2c5}|\x{2c8}|\x{2cc}-\x{2d8}|\x{2da}-\x{2db}|\x{2dd}-\x{377}|\x{37a}-\x{37f}|\x{384}-\x{38a}|\x{38c}|\x{38e}-\x{390}|\x{3aa}-\x{3b0}|\x{3c2}|\x{3ca}-\x{400}|\x{402}-\x{40f}|\x{450}|\x{452}-\x{52f}|\x{531}-\x{556}|\x{559}-\x{55f}|\x{561}-\x{587}|\x{589}-\x{58a}|\x{58d}-\x{58f}|\x{e3f}|\x{1d00}-\x{1dca}|\x{1dfe}-\x{1f15}|\x{1f18}-\x{1f1d}|\x{1f20}-\x{1f45}|\x{1f48}-\x{1f4d}|\x{1f50}-\x{1f57}|\x{1f59}|\x{1f5b}|\x{1f5d}|\x{1f5f}-\x{1f7d}|\x{1f80}-\x{1fb4}|\x{1fb6}-\x{1fc4}|\x{1fc6}-\x{1fd3}|\x{1fd6}-\x{1fdb}|\x{1fdd}-\x{1fef}|\x{1ff2}-\x{1ff4}|\x{1ff6}-\x{1ffe}|\x{2000}-\x{200f}|\x{2011}-\x{2012}|\x{2017}|\x{201b}|\x{201f}|\x{2024}|\x{202f}|\x{2034}|\x{203c}-\x{203e}|\x{2044}-\x{2046}|\x{2057}|\x{205e}-\x{205f}|\x{2061}-\x{2063}|\x{2070}-\x{2071}|\x{2074}-\x{208e}|\x{2090}-\x{209c}|\x{20a0}-\x{20ab}|\x{20ad}-\x{20b5}|\x{20b8}-\x{20ba}|\x{20bc}-\x{20bf}|\x{20d0}-\x{20df}|\x{20e1}|\x{20e5}-\x{20e6}|\x{20e8}-\x{20ea}|\x{2100}-\x{2102}|\x{2104}|\x{2106}-\x{2108}|\x{210a}-\x{2115}|\x{2117}-\x{2120}|\x{2123}-\x{214f}|\x{2153}-\x{215e}|\x{2183}-\x{2184}|\x{2194}-\x{2195}|\x{219a}-\x{2207}|\x{2209}-\x{220e}|\x{2210}|\x{2212}-\x{2214}|\x{2216}-\x{2219}|\x{221b}-\x{221c}|\x{2221}-\x{2222}|\x{2224}|\x{2226}|\x{222c}-\x{222d}|\x{222f}-\x{2233}|\x{2238}-\x{223c}|\x{223e}-\x{2247}|\x{2249}-\x{224b}|\x{224d}-\x{2251}|\x{2253}-\x{225f}|\x{2262}-\x{2263}|\x{2268}-\x{226d}|\x{2270}-\x{2294}|\x{2296}-\x{2298}|\x{229a}-\x{22a4}|\x{22a6}-\x{22be}|\x{22c0}-\x{2311}|\x{2313}-\x{232a}|\x{2330}-\x{23cf}|\x{23dc}-\x{23e0}|\x{246a}-\x{2473}|\x{24ea}-\x{24f4}|\x{24ff}|\x{2592}|\x{25a2}-\x{25b1}|\x{25b4}-\x{25bb}|\x{25be}-\x{25c5}|\x{25c8}-\x{25ca}|\x{25cc}-\x{25cd}|\x{25d0}-\x{25e1}|\x{25e6}-\x{25ff}|\x{2660}-\x{2663}|\x{2666}|\x{2720}|\x{2776}-\x{277f}|\x{27c0}-\x{27ff}|\x{2900}-\x{2aff}|\x{2b04}|\x{2b06}-\x{2b07}|\x{2b0c}-\x{2b0d}|\x{2b1a}|\x{2c60}-\x{2c7f}|\x{2e17}|\x{a64c}-\x{a64d}|\x{a717}-\x{a71a}|\x{a720}-\x{a721}|\x{fb00}-\x{fb04}|\x{fb13}-\x{fb17}|\x{fe00}|\x{1d400}-\x{1d454}|\x{1d456}-\x{1d49c}|\x{1d49e}-\x{1d49f}|\x{1d4a2}|\x{1d4a5}-\x{1d4a6}|\x{1d4a9}-\x{1d4ac}|\x{1d4ae}-\x{1d4b9}|\x{1d4bb}|\x{1d4bd}-\x{1d4c0}|\x{1d4c2}-\x{1d4c3}|\x{1d4c5}-\x{1d505}|\x{1d507}-\x{1d50a}|\x{1d50d}-\x{1d514}|\x{1d516}-\x{1d51c}|\x{1d51e}-\x{1d539}|\x{1d53b}-\x{1d53e}|\x{1d540}-\x{1d544}|\x{1d546}|\x{1d54a}-\x{1d550}|\x{1d552}-\x{1d6a5}|\x{1d6a8}-\x{1d7cb}|\x{1d7ce}-\x{1d7ff}|\x{1d4c1}]+)/u',
        ], [
                '<w:rPr><w:rFonts w:ascii="SimSun-ExtB" w:eastAsia="SimSun-ExtB" w:hAnsi="SimSun-ExtB" w:cs="SimSun-ExtB" w:hint="eastAsia"/></w:rPr>$1'.$font,
                '<w:rPr><w:rFonts w:ascii="Cambria Math" w:eastAsia="Cambria Math" w:hAnsi="Cambria Math" w:cs="Cambria Math" w:hint="eastAsia"/></w:rPr>$1'.$font,
            ], $xml);
    }
    
    private function deleteComments($remainComments = true) {
        if($remainComments) {
            $edit = $this->commentsEdit[0];
            $willDeleted = [];
            $usedCommentIds = $this->wordProcessor->getUsedCommentIds();
            $edit->treeToListCallback($edit->DOMDocument,function($node) use($edit,$usedCommentIds,&$willDeleted) {
                if($node->localName == 'comment' && isset($usedCommentIds[$edit->getAttr($node,'id')])) {
                    $willDeleted[] = $node;
                }else{
                    return $node;
                }
            });

            foreach($willDeleted as $node) {
                $edit->removeChild($node);
            }
        }else{
            $parts = [];
            if(isset($this->parts[15])) {
                $parts = array_merge($parts,$this->parts[15]);
                unset($this->parts[15]);
            }
            if(isset($this->parts[16])) {
                $parts = array_merge($parts,$this->parts[16]);
                unset($this->parts[16]);
            }
            if(isset($this->parts[17])) {
                $parts = array_merge($parts,$this->parts[17]);
                unset($this->parts[17]);
            }
            foreach($parts as $part) {
                $this->zip->deleteName($part['PartName']);
            }
        }

        foreach($this->needUpdateParts as $part) {
            $func = $part['func'];
            switch ($func) {
                case 'getHeaderEdit':
                    $this->deleteMded($this->headerEdits[$part['partName']],$remainComments);
                    break;
                case 'getDocumentEdit':
                    $this->deleteMded($this->wordProcessor->getDocumentEdit(),$remainComments);
                    break;
                case 'getFooterEdit':
                    $this->deleteMded($this->footerEdits[$part['partName']],$remainComments);
                    break;
            }
        }
        
        //test
//         echo $this->documentEdit->DOMDocument->saveXML();exit;
    }
    
    private function deleteMded($edit,$remainComments=false) {
        $deleteTags = [
            'commentRangeStart'=>1,
            'commentRangeEnd'=>1,
            'commentReference'=>1
        ];
        
        $willDeleted = [];
        $usedCommentIds = $this->wordProcessor->getUsedCommentIds();
        if($remainComments) {
            $edit->treeToListCallback($edit->DOMDocument,function($node) use($edit,$deleteTags,$usedCommentIds,&$willDeleted) {
                if($edit->getAttr($node,'md',null)) {
                    $willDeleted[] = $node;
                }else{
                    $id = $edit->getAttr($node,'id');
                    if(isset($deleteTags[$node->localName]) && $id[0] !== 'r' && isset($usedCommentIds[$id])) {
                        $willDeleted[] = $node;
                    }else{
                        return $node;
                    }
                }
            });
        }else{
            $edit->treeToListCallback($edit->DOMDocument,function($node) use($edit,$deleteTags,$usedCommentIds,&$willDeleted) {
                if($edit->getAttr($node,'md',null) || isset($deleteTags[$node->localName])) {
                    $willDeleted[] = $node;
                }else{
                    return $node;
                }
            });
        }
            
        foreach($willDeleted as $node) {
            $edit->removeChild($node);
        }
    }
    
    /**
     *
     * @param string $filename
     * @return \DOMDocument
     */
    public function getXmlDom($filename,$standardXmlFunc=null) {
        if(is_null($filename) && !is_null($standardXmlFunc)) {
            $xml = $standardXmlFunc();
        }else{
            $xml = $this->zip->getFromName($filename);
            if(!is_null($standardXmlFunc)) {
                $xml = $standardXmlFunc($xml);
            }
        }
        $domDocument = new \DOMDocument();
        $domDocument->loadXML($xml);
        if(MDWORD_DEBUG === true) {
            $domDocument->formatOutput = true;
        }
        
        return $domDocument;
    }
    
    protected function getZipFiles() {
        static $pathIndx = null;
        
        if(is_null($pathIndx)) {
            for($i=0;$i < $this->zip->numFiles; $i++) {
                $name = $this->zip->getNameIndex($i);
                $nameArr = explode('/', $name);
                
                $item = &$pathIndx;
                foreach ($nameArr as $pathOrName) {
                    if(!isset($item[$pathOrName])) {
                        $item[$pathOrName] = [];
                    }
                    $item = &$item[$pathOrName];
                }
            }
            unset($item);
        }
        
        return $pathIndx;
    }
   
    public function getChartParts(){
        if(isset($this->parts[13])){
            $chartPartsArr[13] = $this->getPartsDetail($this->parts[13]);
        }
        if(isset($this->parts[25])){
            $chartPartsArr[25] = $this->parts[25];
        }
        if(isset($this->parts[26])){
            $chartPartsArr[26] = $this->parts[26];
        }
        return $chartPartsArr;
    }

    public function getPartsDetail($parts){
        foreach($parts as &$part){
            $partInfo = pathinfo($part['PartName']);
            $partNameRel = $partInfo['dirname'].'/_rels/'.$partInfo['basename'].'.rels';
            $chartRelDom = $this->getXmlDom($partNameRel);
            if(!empty($chartRelDom)){
                $part['chartRelDom'] = $chartRelDom;         
            }
            $Relationships = $chartRelDom->getElementsByTagName('Relationship');
            foreach($Relationships as $Relationship){
                $target = $Relationship->getAttribute('Target');
                if(strpos($target,'xlsx') != false){
                    $part['embeddings']['xml'] = $this->zip->getFromName(str_replace('../','word/',$target));
                    preg_match('/embeddings\/([\s\S]+)/',$target,$match);
                    $part['embeddings']['name'] = $match[1];
                }
            }
        }
        return $parts;
    }

    public function setChartParts($chartparts){
        foreach($chartparts as $key=>&$part){
            if($key==13){
                if(!empty($this->parts[13])){
                    foreach($this->parts[13] as $keypart){
                        $partName = $keypart['PartName'];
                        preg_match('/(\d+)/',$partName,$match);
                        $chartKey = $match[1];
                    }
                    foreach($part as &$p){
                        $chartKey++;
                        preg_match('/(\d+)/',$p['PartName'],$partmatch);
                        $p['PartName'] = str_replace($partmatch[1],$chartKey,$p['PartName']);
                        $styleKey[$partmatch[1]] = $chartKey;              
                    }
                    foreach($chartparts[25] as &$stylepart){
                        preg_match('/(\d+)/',$stylepart['PartName'],$styleMatch);
                        $stylepart['PartName'] = str_replace($styleMatch[1],$styleKey[$styleMatch[1]],$stylepart['PartName']);
                    }   
                    foreach($chartparts[26] as &$colorepart){
                        preg_match('/(\d+)/',$colorepart['PartName'],$colorMatch);
                        $colorepart['PartName'] = str_replace($colorMatch[1],$styleKey[$colorMatch[1]],$colorepart['PartName']);
                    }   
                    $this->parts[13] = array_merge($this->parts[13],$chartparts[13]);
                }else{
                    $this->parts[13] = $chartparts[13];
                }
            }else{
                $this->parts[$key] = empty($this->parts[$key])?$chartparts[$key]:array_merge($this->parts[$key],$chartparts[$key]);
            }
        }
    }

    public function getContentTypes(){
        $contentTypes = $this->Content_Types->overrides;

        return $contentTypes;
    }

    public function getChartEmbeddings(){
        $fileList = $this->getZipFiles();
        $relArr = $fileList['word']['charts']['_rels'];
        $embeddingsList = $fileList['word']['embeddings'];
        //$this->documentEdit->setChartRelValue($relArr);
        foreach($embeddingsList as $embeddingKey=>$val){
            $embeddingsXml[$embeddingKey] = $this->zip->getFromName('word/embeddings/'.$embeddingKey);
        }
        return $embeddingsXml;
    }

    public function setContentTypes(){
        foreach($this->parts[13] as $part){
            $partArr['PartName'] = $part['PartName'];
            $partArr['ContentType'] = 13;
            if(array_search($partArr,$this->Content_Types->overrides) === false){
                $newOverrides[] = $partArr;
            }
        }
        foreach($this->parts[25] as $part){
            $partArr['PartName'] = $part['PartName'];
            $partArr['ContentType'] = 25;
            if(array_search($partArr,$this->Content_Types->overrides) === false){
                $newOverrides[] = $partArr;
            }
        }
        foreach($this->parts[26] as $part){
            $partArr['PartName'] = $part['PartName'];
            $partArr['ContentType'] = 26;
            if(array_search($partArr,$this->Content_Types->overrides) === false){
                $newOverrides[] = $partArr;
            }
        }
        $this->Content_Types->setContent_types($newOverrides);
        $this->writeContentTypeXml();

    }

    public function writeContentTypeXml(){
        $Relationships = $this->Content_Types->DOMDocument->getElementsByTagName('Override');
        foreach($this->Content_Types->overrides as $key=>$overrides){
            if(!isset($Relationships[$key])){
                $copy = clone $Relationships[0];
                $copy->setAttribute('PartName','/'.$overrides['PartName']);
                $copy->setAttribute('ContentType',$this->Content_Types->contentTypes[$overrides['ContentType']]);
                $Relationships[0]->parentNode->appendChild($copy);
            }
        }
        $this->zip->addFromString('[Content_Types].xml', $this->Content_Types->DOMDocument->saveXML());
    }

    
    public function updateChartRel(){
        foreach($this->parts[13] as $part){
            if(isset($part['chartRelDom'])){
                $Relationships = $part['chartRelDom']->getElementsByTagName('Relationship');
                preg_match('/chart(\d+)/',$part['PartName'],$match);
                foreach($Relationships as $Relationship){
                    $target = $Relationship->getAttribute('Target');
                    $Relationship->setAttribute('Target',preg_replace('/(\d+)/',$match[1],$target));
                }
                $relArr['PartName'] =  $part['PartName'];
                $relArr['relName'] =  'word/charts/_rels/'.$match[0].'.xml.rels';
                $relArr['dom'] =  $part['chartRelDom'];

                $this->documentEdit->setChartRel($relArr);
            }
        }
    }

    public function free() {
        unset($this->Content_Types->word);
        unset($this->Content_Types);

        unset($this->zip,$this->log);
        $this->documentEdit->free();

        unset($this->documentEdit->DOMDocument);
        unset($this->documentEdit->commentsEdit);
        
        unset($this->commentsEdit);
        unset($this->stylesEdit);
        
        //fixed bug:not use unset
        // unset($this->documentEdit);
        $this->documentEdit = null;
    }
}
