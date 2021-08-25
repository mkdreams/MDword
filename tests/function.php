<?php
function markdownTable($datas) {
    $tableStr = '';
    $count = count($datas[0]);
    $lines = [];
    $line = '';
    $splitLine = '';
    foreach($datas[0] as $str) {
        $line .= '|  '.$str.'   ';
        $splitLine .= '|  ----   ';
    }
    $line .= '|';
    $splitLine .= '|';

    $lines[] = $line;
    $lines[] = $splitLine;

    foreach ($datas as $k => $v) {
        if($k === 0) {
            continue;
        }

        $line = '';
        foreach($v as $str) {
            $line .= '|  '.$str.'   ';
        }
        $line .= '|';

        $lines[] = $line;
    }

    return implode("\r\n",$lines);
}
