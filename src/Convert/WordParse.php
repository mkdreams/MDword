<?php 
namespace MDword\Convert;

use MDword\Common\Build;

class WordParse{
    private $defualtStyle = 
    //--WORDPARSE-DEFUALTSTYLE--
array (
  'ParaPr' => 
  array (
    'ContextualSpacing' => false,
    'Ind' => 
    array (
      'Left' => 0,
      'Right' => 0,
      'FirstLine' => 0,
    ),
    'Jc' => 1,
    'KeepLines' => false,
    'KeepNext' => false,
    'PageBreakBefore' => false,
    'Spacing' => 
    array (
      'Line' => 1.1499999999999999,
      'LineRule' => 1,
      'Before' => 0,
      'BeforeAutoSpacing' => false,
      'After' => 3.5277777777777777,
      'AfterAutoSpacing' => false,
    ),
    'Shd' => 
    array (
      'Value' => 1,
      'Color' => 
      array (
        'r' => 255,
        'g' => 255,
        'b' => 255,
        'Auto' => false,
      ),
    ),
    'Brd' => 
    array (
      'First' => true,
      'Last' => true,
      'Between' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
      'Bottom' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
      'Left' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
      'Right' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
      'Top' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
    ),
    'WidowControl' => true,
    'Tabs' => 
    array (
      'Tabs' => 
      array (
      ),
    ),
  ),
  'TextPr' => 
  array (
    'Bold' => false,
    'Italic' => false,
    'Strikeout' => false,
    'Underline' => false,
    'FontFamily' => 
    array (
      'Name' => 'Arial',
      'Index' => -1,
    ),
    'FontSize' => 11,
    'Color' => 
    array (
      'r' => 0,
      'g' => 0,
      'b' => 0,
      'Auto' => true,
    ),
    'VertAlign' => 0,
    'HighLight' => -1,
    'Spacing' => 0,
    'DStrikeout' => false,
    'Caps' => false,
    'SmallCaps' => false,
    'Position' => 0,
    'RFonts' => 
    array (
      'Ascii' => 
      array (
        'Name' => 'Arial',
        'Index' => -1,
      ),
      'EastAsia' => 
      array (
        'Name' => 'Arial',
        'Index' => -1,
      ),
      'HAnsi' => 
      array (
        'Name' => 'Arial',
        'Index' => -1,
      ),
      'CS' => 
      array (
        'Name' => 'Arial',
        'Index' => -1,
      ),
      'Hint' => 0,
    ),
    'BoldCS' => false,
    'ItalicCS' => false,
    'FontSizeCS' => 11,
    'CS' => false,
    'RTL' => false,
    'Lang' => 
    array (
      'Bidi' => 1033,
      'EastAsia' => 1033,
      'Val' => 1033,
    ),
    'Vanish' => false,
  ),
  'TablePr' => 
  array (
    'TableStyleColBandSize' => 1,
    'TableStyleRowBandSize' => 1,
    'Jc' => 1,
    'Shd' => 
    array (
      'Value' => 1,
      'Color' => 
      array (
        'r' => 255,
        'g' => 255,
        'b' => 255,
        'Auto' => false,
      ),
    ),
    'TableBorders' => 
    array (
      'Bottom' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
      'Left' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
      'Right' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
      'Top' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
      'InsideH' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
      'InsideV' => 
      array (
        'Color' => 
        array (
          'r' => 0,
          'g' => 0,
          'b' => 0,
          'Auto' => false,
        ),
        'Space' => 0,
        'Size' => 0.17638888888888887,
        'Value' => 0,
      ),
    ),
    'TableCellMar' => 
    array (
      'Bottom' => 
      array (
        'Type' => 1,
        'W' => 0,
      ),
      'Left' => 
      array (
        'Type' => 1,
        'W' => 0.17638888888888887,
      ),
      'Right' => 
      array (
        'Type' => 1,
        'W' => 0.17638888888888887,
      ),
      'Top' => 
      array (
        'Type' => 1,
        'W' => 0,
      ),
    ),
    'TableCellSpacing' => NULL,
    'TableInd' => 0,
    'TableW' => 
    array (
      'Type' => 0,
      'W' => 0,
    ),
    'TableLayout' => 1,
    'TableDescription' => '',
    'TableCaption' => '',
  ),
  'TableRowPr' => 
  array (
    'CantSplit' => false,
    'GridAfter' => 0,
    'GridBefore' => 0,
    'Jc' => 1,
    'TableCellSpacing' => NULL,
    'Height' => 
    array (
      'Value' => 0,
      'HRule' => 1,
    ),
    'WAfter' => 
    array (
      'Type' => 0,
      'W' => 0,
    ),
    'WBefore' => 
    array (
      'Type' => 0,
      'W' => 0,
    ),
    'TableHeader' => false,
  ),
  'TableCellPr' => 
  array (
    'GridSpan' => 1,
    'Shd' => 
    array (
      'Value' => 1,
      'Color' => 
      array (
        'r' => 255,
        'g' => 255,
        'b' => 255,
        'Auto' => false,
      ),
    ),
    'TableCellMar' => NULL,
    'TableCellBorders' => 
    array (
    ),
    'TableCellW' => 
    array (
      'Type' => 0,
      'W' => 0,
    ),
    'VAlign' => 0,
    'VMerge' => 1,
    'TextDirection' => 0,
    'NoWrap' => false,
    'HMerge' => 1,
  ),
  'Paragraph' => NULL,
  'Character' => NULL,
  'Numbering' => NULL,
  'Table' => NULL,
  'TableGrid' => '46',
  'Headings' => 
  array (
    0 => '11',
    1 => '13',
    2 => '15',
    3 => '17',
    4 => '19',
    5 => '21',
    6 => '23',
    7 => '25',
    8 => '27',
  ),
  'ParaList' => '29',
  'Header' => '40',
  'Footer' => '42',
  'Hyperlink' => '172',
  'FootnoteText' => '173',
  'FootnoteTextChar' => '174',
  'FootnoteReference' => '175',
  'NoSpacing' => '31',
  'Title' => '32',
  'Subtitle' => '34',
  'Quote' => '36',
  'IntenseQuote' => '38',
  'TOC' => 
  array (
    0 => '176',
    1 => '177',
    2 => '178',
    3 => '179',
    4 => '180',
    5 => '181',
    6 => '182',
    7 => '183',
    8 => '184',
  ),
  'TOCHeading' => '185',
  'Caption' => '44',
)//--WORDPARSE-DEFUALTSTYLE--
    ;
    private $tableTypes = 
    //--WORDPARSE-TABLETYPES--
array (
  'Signature' => 0,
  'Info' => 1,
  'Media' => 2,
  'Numbering' => 3,
  'HdrFtr' => 4,
  'Style' => 5,
  'Document' => 6,
  'Other' => 7,
  'Comments' => 8,
  'Settings' => 9,
  'Footnotes' => 10,
  'Endnotes' => 11,
  'Background' => 12,
  'VbaProject' => 13,
  'App' => 15,
  'Core' => 16,
  'DocumentComments' => 17,
)//--WORDPARSE-TABLETYPES--
    ;
    private $stream;
    public $versionInfo = [];
    /**
     * @param Stream $stream
     */
    public function __construct($stream) {
        $this->stream = $stream;
    }
    
    private function initStyle() {
        $this->versionInfo = $this->stream->getVersionInfo();
        
        if(MDWORD_DEBUG) {
            $build = new Build();
            
            $this->defualtStyle = json_decode('{"ParaPr":{"ContextualSpacing":false,"Ind":{"Left":0,"Right":0,"FirstLine":0},"Jc":1,"KeepLines":false,"KeepNext":false,"PageBreakBefore":false,"Spacing":{"Line":1.15,"LineRule":1,"Before":0,"BeforeAutoSpacing":false,"After":3.5277777777777777,"AfterAutoSpacing":false},"Shd":{"Value":1,"Color":{"r":255,"g":255,"b":255,"Auto":false}},"Brd":{"First":true,"Last":true,"Between":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0},"Bottom":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0},"Left":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0},"Right":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0},"Top":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0}},"WidowControl":true,"Tabs":{"Tabs":[]}},"TextPr":{"Bold":false,"Italic":false,"Strikeout":false,"Underline":false,"FontFamily":{"Name":"Arial","Index":-1},"FontSize":11,"Color":{"r":0,"g":0,"b":0,"Auto":true},"VertAlign":0,"HighLight":-1,"Spacing":0,"DStrikeout":false,"Caps":false,"SmallCaps":false,"Position":0,"RFonts":{"Ascii":{"Name":"Arial","Index":-1},"EastAsia":{"Name":"Arial","Index":-1},"HAnsi":{"Name":"Arial","Index":-1},"CS":{"Name":"Arial","Index":-1},"Hint":0},"BoldCS":false,"ItalicCS":false,"FontSizeCS":11,"CS":false,"RTL":false,"Lang":{"Bidi":1033,"EastAsia":1033,"Val":1033},"Vanish":false},"TablePr":{"TableStyleColBandSize":1,"TableStyleRowBandSize":1,"Jc":1,"Shd":{"Value":1,"Color":{"r":255,"g":255,"b":255,"Auto":false}},"TableBorders":{"Bottom":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0},"Left":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0},"Right":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0},"Top":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0},"InsideH":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0},"InsideV":{"Color":{"r":0,"g":0,"b":0,"Auto":false},"Space":0,"Size":0.17638888888888887,"Value":0}},"TableCellMar":{"Bottom":{"Type":1,"W":0},"Left":{"Type":1,"W":0.17638888888888887},"Right":{"Type":1,"W":0.17638888888888887},"Top":{"Type":1,"W":0}},"TableCellSpacing":null,"TableInd":0,"TableW":{"Type":0,"W":0},"TableLayout":1,"TableDescription":"","TableCaption":""},"TableRowPr":{"CantSplit":false,"GridAfter":0,"GridBefore":0,"Jc":1,"TableCellSpacing":null,"Height":{"Value":0,"HRule":1},"WAfter":{"Type":0,"W":0},"WBefore":{"Type":0,"W":0},"TableHeader":false},"TableCellPr":{"GridSpan":1,"Shd":{"Value":1,"Color":{"r":255,"g":255,"b":255,"Auto":false}},"TableCellMar":null,"TableCellBorders":{},"TableCellW":{"Type":0,"W":0},"VAlign":0,"VMerge":1,"TextDirection":0,"NoWrap":false,"HMerge":1},"Paragraph":"8","Character":"9","Numbering":"10","Table":"30","TableGrid":"46","Headings":["11","13","15","17","19","21","23","25","27"],"ParaList":"29","Header":"40","Footer":"42","Hyperlink":"172","FootnoteText":"173","FootnoteTextChar":"174","FootnoteReference":"175","NoSpacing":"31","Title":"32","Subtitle":"34","Quote":"36","IntenseQuote":"38","TOC":["176","177","178","179","180","181","182","183","184"],"TOCHeading":"185","Caption":"44"}',true);
            $this->defualtStyle['Character'] = null;
            $this->defualtStyle['Numbering'] = null;
            $this->defualtStyle['Paragraph'] = null;
            $this->defualtStyle['Table'] = null;
            
            $build->replace('WORDPARSE-DEFUALTSTYLE', $this->defualtStyle, __FILE__);
            
            $this->tableTypes = json_decode('{"Signature":0,"Info":1,"Media":2,"Numbering":3,"HdrFtr":4,"Style":5,"Document":6,"Other":7,"Comments":8,"Settings":9,"Footnotes":10,"Endnotes":11,"Background":12,"VbaProject":13,"App":15,"Core":16,"DocumentComments":17}',true);
            
            $build->replace('WORDPARSE-TABLETYPES', $this->tableTypes, __FILE__);
        }
        
    }
    
    public function readBin() {
        $this->initStyle();
        $this->readMainTable();
    }
    
    public function readMainTable() {
        $this->stream->enterFrame(1);
        
        $mtLen = $this->stream->getUChar();
        
        $aSeekTable = [];
        $tableSeeks = [];
        $tableSeeks['nOtherTableSeek'] = -1;
        $tableSeeks['nNumberingTableSeek'] = -1;
        $tableSeeks['nCommentTableSeek'] = -1;
        $tableSeeks['nDocumentCommentTableSeek'] = -1;
        $tableSeeks['nSettingTableSeek'] = -1;
        $tableSeeks['nDocumentTableSeek'] = -1;
        $tableSeeks['nFootnoteTableSeek'] = -1;
        
//         $fileStream;
        for($i = 0; $i < $mtLen; $i++) {
            $this->stream->enterFrame(5);//type + len=>1 + 4
            $mtiType = $this->stream->getUChar();
            $mtiOffBits = $this->stream->getULongLE();
            
            switch ($mtiType) {
                case $this->tableTypes['Other']:
                    $tableSeeks['nOtherTableSeek'] = $mtiOffBits;
                    break;
                case $this->tableTypes['Numbering']:
                    $tableSeeks['nNumberingTableSeek'] = $mtiOffBits;
                    break;
                case $this->tableTypes['Comments']:
                    $tableSeeks['nCommentTableSeek'] = $mtiOffBits;
                    break;
                case $this->tableTypes['DocumentComments']:
                    $tableSeeks['nDocumentCommentTableSeek'] = $mtiOffBits;
                    break;
                case $this->tableTypes['Settings']:
                    $tableSeeks['nSettingTableSeek'] = $mtiOffBits;
                    break;
                case $this->tableTypes['Document']:
                    $tableSeeks['nDocumentTableSeek'] = $mtiOffBits;
                    break;
                case $this->tableTypes['Footnotes']:
                    $tableSeeks['nFootnoteTableSeek'] = $mtiOffBits;
                    break;
                default:
                    $aSeekTable[] = ['type' => $mtiType, 'offset' => $mtiOffBits];
                    break;
            }
        }
        
        if($tableSeeks['nOtherTableSeek'] > 0) {
            $this->stream->seek($tableSeeks['nOtherTableSeek']);
        }
        
        var_dump($aSeekTable);exit;
        
        
        var_dump($mtLen);
    }
}