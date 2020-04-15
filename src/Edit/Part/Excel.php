<?php
namespace MDword\Edit\Part;

use MDword\Common\PartBase;

class Excel extends PartBase
{
    /**
     * phpexcel use in chart
     *
     * @var \PHPExcel_Reader_Excel2007
     */
    private $spreadsheet = null;
    
    private $PHPExcel_Writer = null;
    public function __construct($word,$partName) {
        $this->word = $word;
        
        if(class_exists('PhpOffice\PhpSpreadsheet\Spreadsheet')) {
            $this->spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        }else{
            $this->spreadsheet = new \PHPExcel_Reader_Excel2007();
        }
        
        $partName = $this->pathRelToAbs($partName);
        $content = $this->word->zip->getFromName($partName);
        $this->tempDocumentFilename = tempnam($this->getTempDir(), 'MDwordExcelPart');
        file_put_contents($this->tempDocumentFilename, $content);
        
        $this->word->files[] = ['type'=>'excel','PartName'=>$partName,'this'=>$this];
        
        $this->spreadsheet = $this->spreadsheet->load($this->tempDocumentFilename);
    }

    public function changeExcelValues($datas) {
        foreach($datas as $data) {
            if($data[2] === 'set') {
                $currentSheet = $this->spreadsheet->getSheetByName($data[0][0]);
                $currentSheet->setCellValue($data[0][1][0].$data[0][2][0],$data[1]);
            }
        }
    }
    
    public function preDealDatas ($datas) {
        $setDatas = [];
        $extDatas = [];
        foreach($datas as $data) {
            if($data[2] == 'set') {
                $data[0] = $this->parserRange($data[0]);
                $setDatas[] = $data;
            }else if($data[2] == 'ext') {
                $data[0] = $this->parserRange($data[0]);
                $data[1] = $this->parserRange($data[1]);
                $extDatas[] = $data;
            }
        }
        
        return array_merge($extDatas,$setDatas);
    }
    
    public function getContent() {
        $this->PHPExcel_Writer = new \PHPExcel_Writer_Excel2007($this->spreadsheet);
        $this->PHPExcel_Writer->save($this->tempDocumentFilename);
        $content = file_get_contents($this->tempDocumentFilename);
        @unlink($this->tempDocumentFilename);
        return $content;
    }
}
