<?php
namespace MDword\Api;
require_once(__DIR__.DIRECTORY_SEPARATOR.'..'.DIRECTORY_SEPARATOR.'Config'.DIRECTORY_SEPARATOR.'Main.php');

use MDword\WordProcessor;
use MDword\Common\Common;

class Base{
    protected $parameters = [];
    protected $common = null;
    protected $wordProcessor = null;
    public function __construct() {
        $this->common = new Common();
        
        $this->parseParameters();
        $this->wordProcessor = new WordProcessor();
        $this->wordProcessor->load($this->parameters['docUrl']);
    }
    
    private function parseParameters() {
        $this->parameters = [
            'docUrl'=>MDWORD_TEST_DIRECTORY.'/samples/api/temple.docx',
            'datas'=>[
                'data1'=>[
                    'type'=>'json',
                    'dataInfo'=>[
                        'type'=>'local',
                        'data'=>json_encode([
                            ['price'=>100,'change'=>5,'changepercent'=>0.05],
                            ['price'=>200,'change'=>-10,'changepercent'=>-0.05],
                            ['price'=>500,'change'=>100,'changepercent'=>0.20],
                        ])
                    ],
                ],
                'data2'=>[
                    'type'=>'json',
                    'dataInfo'=>[
                        'type'=>'http',
                        'url'=>'http://restapi2.farseerbi.com/bds/ipoNews?accessToken=3eZF3JrHUtbd5kXc1CVVrLAQf7XUJVWs',
                        'headers'=>'',
                        'postData'=>['rows'=>10,'sLanguage'=>'TC','page'=>1],
                    ],
                ],
                'data3'=>[
                    'type'=>'csv',
                    'dataInfo'=>[
                        'type'=>'http',
                        'url'=>MDWORD_TEST_DIRECTORY.'/samples/api/simple data.csv',
                        'headers'=>'',
                        'postData'=>'',
                    ],
                ],
                'data4'=>[
                    'type'=>'xml'
                ],
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
                'rows'=>[
                    'dataKeyName'=>'data2',
                    'keyList'=>['rows'],
                    'childrens'=>[
                        'title'=>[
                            'keyList'=>['newsSubj'],
                        ],
                        'company name'=>[
                            'keyList'=>['stockName'],
                        ],
                        'date'=>[
                            'keyList'=>['newsDate'],
                        ],
                    ]
                ],
                'rows2'=>[
                    'dataKeyName'=>'data3',
                    'keyList'=>[],
                    'childrens'=>[
                        'type2'=>[
                            'keyList'=>[0],
                        ],
                        'title2'=>[
                            'keyList'=>[1],
                        ],
                        'influence2'=>[
                            'keyList'=>[6],
                        ],
                        'date2'=>[
                            'keyList'=>[3],
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
