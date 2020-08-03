<?php

MD_Autoloader::Register();

class MD_Autoloader
{
    /**
     * Register the Autoloader with SPL
     *
     */
    public static function Register() {
        if (function_exists('__autoload')) {
            //    Register any existing autoloader function with SPL, so we don't get any clashes
            spl_autoload_register('__autoload');
        }
        //    Register ourselves with SPL
        return spl_autoload_register(array('MD_Autoloader', 'Load'));
    }   //    function Register()


    /**
     * Autoload a class identified by name
     *
     * @param    string    $pClassName        Name of the object to load
     */
    public static function Load($pClassName){
        if ((class_exists($pClassName,FALSE)) || (strpos($pClassName, 'MDword') !== 0)) {
            //    Either already loaded, or not a PHPExcel class request
            return FALSE;
        }

        $pClassNameArr = explode('\\', $pClassName,2);
        
        $pClassNameArr[1] = str_replace('\\', DIRECTORY_SEPARATOR,$pClassNameArr[1]);
        
        $pClassFilePath = dirname(__FILE__).'/src/'.
            $pClassNameArr[1].
                          '.php';

        if ((file_exists($pClassFilePath) === FALSE) || (is_readable($pClassFilePath) === FALSE)) {
            //    Can't load
            return FALSE;
        }

        require($pClassFilePath);
    }   //    function Load()

}
