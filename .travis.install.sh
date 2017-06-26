#!/bin/sh
set -ev
if [ "$PHP_LATEST_VERSION" != 1 ]
then
    composer remove --no-update --no-scripts --dev slam/php-cs-fixer-extensions
fi
composer install --prefer-dist
