<?php

$config = new SlamCsFixer\Config();
$config->getFinder()
    ->in(__DIR__ . '/lib/Exception')
    ->in(__DIR__ . '/lib/Helper')
    ->in(__DIR__ . '/tests')
;

return $config;
