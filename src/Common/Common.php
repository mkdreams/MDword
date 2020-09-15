<?php
namespace MDword\Common;

class Common
{
    public $log = null;
    public function __construct() {
        $this->log = new Log();
    }
    
    public function getDirFiles($dir,$callback=null) {
        if(!is_dir($dir)) {
            $this->log->writeLog('dir is not a dir! dir:'.$dir);
            return [];
        }
        $filesTemp = scandir($dir);
        $files = [];
        foreach ($filesTemp as $fileName){
            if($fileName != '.' && $fileName != '..'){
                if(!is_null($callback)) {
                    $fileName = $callback($dir,$fileName);
                }
                $files[] = $fileName;
            }
        }
        
        
        return $files;
    }
    
    public function CurlSend($url,$headers='',$post=[],$timeoutMs=30000) {
        $ch = curl_init();
        curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, FALSE);//https
        curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, FALSE);//https
        curl_setopt($ch, CURLOPT_URL, $url);
        curl_setopt($ch, CURLOPT_HEADER, 0);
        curl_setopt($ch, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt($ch, CURLOPT_FOLLOWLOCATION, 1);
        curl_setopt($ch, CURLOPT_TIMEOUT_MS, $timeoutMs);
        
        if(!empty($post)) {
            curl_setopt($ch, CURLOPT_POST, 1);//post方式提交
            if(is_array($post))
                curl_setopt($ch, CURLOPT_POSTFIELDS, http_build_query($post));//要提交的信息
            else
                curl_setopt($ch, CURLOPT_POSTFIELDS, $post);//要提交的信息
                    
        }
        
        $rs = curl_exec($ch); //执行cURL抓取页面内容
        curl_close($ch);
        
        
        return $rs;
    }
}