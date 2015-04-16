<?php

error_reporting(E_ALL);
ini_set('display_errors', true);

set_error_handler(function ($errno, $errstr = '', $errfile = '', $errline = 0) {
    // Controllo necessario per l'operatore @ di soppressione
    if (error_reporting() === 0) {
        return;
    }

    throw new ErrorException($errstr, $errno, $errno, $errfile, $errline);
});

define('ROOT_PATH', dirname(__DIR__));
define('TMP_PATH', ROOT_PATH . '/tmp');

$loader = require ROOT_PATH . '/vendor/autoload.php';
