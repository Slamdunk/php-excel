<?php

$config = new SlamCsFixer\Config(SlamCsFixer\Config::LIB);
$config->setRules([
'no_unneeded_control_parentheses' => true,
]);
$config->getFinder()
    ->in(__DIR__ . '/lib/Pear')
;

return $config;
