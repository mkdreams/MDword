<?php
namespace MDword\Common;

class Build
{
    public function __construct() {
    }
    
    /**
     * must is
     * //--CONTENTTYPES--
     * //--CONTENTTYPES--
     * @param string $name
     * @param string $value
     * @param string $file
     * @throws \Exception
     */
    public function replace($name,$value,$file) {
        if(!is_string($value)) {
            $value = var_export($value,true);
        }
        
        if(!is_file($file)) {
            throw new \Exception('Not a valid file : '.$file);
        }
        
        if(empty($value)) {
            throw new \Exception('value is empty!');
        }
        $content = file_get_contents($file);
        
        $preg = preg_quote($name);
        $content = preg_replace_callback('/(\/\/\-\-'.$preg.'\-\-\s)([\s\S]*)(\/\/\-\-'.$preg.'\-\-)/i', function($match) use($value){
            return $match[1].$value.$match[3];
        }, $content);
        
        $result = file_put_contents($file, $content);
    }
}