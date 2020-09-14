<?php
namespace MDword\Api;
require_once(__DIR__.DIRECTORY_SEPARATOR.'..'.DIRECTORY_SEPARATOR.'Config'.DIRECTORY_SEPARATOR.'Main.php');

use MDword\WordProcessor;

class Base{
    protected $parameters = [];
    protected $wordProcessor = null;
    public function __construct() {
        $this->parseParameters();
        $this->wordProcessor = new WordProcessor();
        $this->wordProcessor->load($this->parameters['docUrl']);
    }
    
    private function parseParameters() {
        $this->parameters = [
            'docUrl'=>MDWORD_TEST_DIRECTORY.'/samples/block/temple.docx',
            'datas'=>[
                'data1'=>[
                    'type'=>'json',
                    'data'=>[
                        ['price'=>100,'change'=>5,'changepercent'=>0.05],
                        ['price'=>200,'change'=>-10,'changepercent'=>-0.05],
                        ['price'=>500,'change'=>100,'changepercent'=>0.20],
                    ]
                ],
                'data2'=>[
                    'type'=>'http',
                    'url'=>'http://47.52.91.57:8081/bds/ipoNews?accessToken=nxNDMvVTu7yfE0MyzmkLJ6MvwEvOaQsn',
                    'headers'=>'',
                    'postData'=>'{"rows":10,"sLanguage":"TC","page":1}',
                ],
                'data3'=>[
                    'type'=>'xml'
                ],
                'data4'=>[
                    'type'=>'excel'
                ],
                'data5'=>[
                    'type'=>'csv'
                ]
            ],
            'binds'=>[
                'item'=>[
                    'dataKeyName'=>'data1',
                    'keyList'=>[],
                    'childrens'=>[
                        'stockprice'=>[
                            'keyList'=>['price'],
                        ],
                        'change'=>[
                            'keyList'=>['change'],
                        ],
                        'changepercent'=>[
                            'keyList'=>['changepercent'],
                        ],
                    ]
                ],
                'two'=>2,
                'box'=>'BOX',
                'header'=>'MDWORD-HEADER',
                'footer'=>'MDWORD-FOOTER',
            ]
        ];
    }
    
    protected function error($text) {
        header('Content-Type:application/json; charset=utf-8');
        $data = ['data'=>null,'messages'=>['message'=>$text,'type'=>'json'],'success'=>false];
        exit(json_encode($data));
    }
    
    protected function success($values,$text='') {
        header('Content-Type:application/json; charset=utf-8');
        $data = ['data'=>$values,'messages'=>['message'=>$text,'type'=>'json'],'success'=>true];
        exit(json_encode($data));
    }
}
