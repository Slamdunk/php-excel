<?php

$config = new SlamCsFixer\Config(false);
$config->getFinder()
    ->in(__DIR__ . '/lib')
    ->in(__DIR__ . '/tests')
;

return $config;
