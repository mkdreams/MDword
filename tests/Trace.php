<?php
use MDword\Common\Build;

require_once(__DIR__.'/../Autoloader.php');
require_once(__DIR__.'/../src/config/main.php');

$build = new Build();

$runSampleFile = realpath(__DIR__.'/samples/simple for readme/index.php');

$dir = dirname($runSampleFile).'/result';
if(!is_dir($dir)) {
    mkdir($dir,'0777',true);
}

$baseName = 'word';

$SAVEANIMALCODE = <<<CODE
\$this->word->wordProcessor->saveAsToPathForTrace('$dir', '$baseName');
// throw new \Exception('saved!');
exit;
CODE;
$build->replace('SAVE-ANIMALCODE', $SAVEANIMALCODE, MDWORD_SRC_DIRECTORY.'/Edit/Part/Document.php');

$maxLine = 0;
try {
    require $runSampleFile;
} catch (Exception $e) {
    $traces = $e->getTrace();
    foreach ($traces as $trace) {
        if(isset($trace['file']) && $trace['file'] === $runSampleFile) {
            $line = $trace['line'];
            if($line > $maxLine) {
                //todo
            }
            var_dump($trace);exit;
        }
    }
}
// while (true) {
// }

// $build->replace('SAVE-ANIMALCODE', '', MDWORD_SRC_DIRECTORY.'/Edit/Part/Document.php');