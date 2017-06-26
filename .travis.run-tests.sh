#!/bin/sh
set -ev
if [ "$PHP_LATEST_VERSION" = 1 ]
then
    composer install --prefer-dist
    vendor/bin/phpunit --coverage-clover ./clover.xml
    phpenv config-rm xdebug.ini || return 0
    vendor/bin/php-cs-fixer fix --diff --dry-run --verbose
    wget https://scrutinizer-ci.com/ocular.phar || return 0
    php ocular.phar code-coverage:upload --format=php-clover ./clover.xml || return 0
else
    phpenv config-rm xdebug.ini || return 0
    composer remove --no-update --no-scripts --dev slam/php-cs-fixer-extensions
    composer install --prefer-dist
    vendor/bin/phpunit
fi
