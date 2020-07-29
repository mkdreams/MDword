<?php 
define('MDWORD_DEBUG', false);
define('MDWORD_GENERATED_DIRECTORY', dirname(__FILE__,3).'/tests/Output');
define('MDWORD_SRC_DIRECTORY', dirname(__FILE__,2));
//1 example: ${name} in comments
//2 example: name in comments
define('MDWORD_BIND_TYPE', 2);

define('MDWORD_BREAK', 1);
define('MDWORD_PAGE_BREAK', 2);
define('MDWORD_TEXT', 3);
define('MDWORD_LINK', 4);
define('MDWORD_IMG', 5);
define('MDWORD_DELETE', 6);