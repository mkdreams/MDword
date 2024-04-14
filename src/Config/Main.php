<?php 
ini_set("pcre.backtrack_limit",10000000);
ini_set("pcre.jit",0);

define('MDWORD_DEBUG', false);
define('MDWORD_TEST_DIRECTORY', dirname(dirname(dirname(__FILE__))).'/tests');
define('MDWORD_GENERATED_DIRECTORY', MDWORD_TEST_DIRECTORY.'/Output');
define('MDWORD_SRC_DIRECTORY', dirname(dirname(__FILE__)));
//1 example: ${name} in comments
//2 example: name in comments
define('MDWORD_BIND_TYPE', 2);

define('MDWORD_BREAK', 1);
define('MDWORD_PAGE_BREAK', 2);
define('MDWORD_TEXT', 3);
define('MDWORD_LINK', 4);
define('MDWORD_IMG', 5);
define('MDWORD_DELETE', 6);
define('MDWORD_PHPWORD', 10);

//update fields: page
define('MDWORD_REF', 12);
define('MDWORD_PAGEREF', 13);
define('MDWORD_NOWPAGE', 14);
define('MDWORD_TOTALPAGE', 15);

define('MDWORD_CLONE', 7);
define('MDWORD_CLONEP', 8);
define('MDWORD_CLONESECTION', 16);
define('MDWORD_CLONETO', 9);
define('MDWORD_TABLE', 11);