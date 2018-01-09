# Slam PHPExcel old&faster

[![Build Status](https://travis-ci.org/Slamdunk/php-excel.svg?branch=master)](https://travis-ci.org/Slamdunk/php-excel)
[![Code Coverage](https://scrutinizer-ci.com/g/Slamdunk/php-excel/badges/coverage.png?b=master)](https://scrutinizer-ci.com/g/Slamdunk/php-excel/?branch=master)
[![Packagist](https://img.shields.io/packagist/v/slam/php-excel.svg)](https://packagist.org/packages/slam/php-excel)

This package is _NOT_ intended to be complete and flexible, but to be *fast*.

[PHPOffice/PHPExcel](https://github.com/PHPOffice/PHPExcel) and [PHPOffice/PhpSpreadsheet](https://github.com/PHPOffice/PhpSpreadsheet) are great libraries,
but abstract everything in memory before writing to the disk. This is
extremely inefficent and slow if you need to write a giant XLS with thousands
rows and hundreds columns.

Based on [Spreadsheet_Excel_Writer v0.9.3](http://pear.php.net/package/Spreadsheet_Excel_Writer),
which can be found active on [Github](https://github.com/pear/Spreadsheet_Excel_Writer).
This is not a fork: I copied it and adapted to work with PHP 7.1 and applied
some coding standard fixes and some Scrutinizer patches.

## Installation

`composer require slam/php-excel`

## Usage

From version 4 the code is split in two parts:

1. `Slam\Excel\Pear` namespace, the original Pear code
1. `Slam\Excel\Helper` namespace, an helper to apply a trivial style on a Table structure:

```php
use Slam\Excel\Helper as ExcelHelper;

require __DIR__ . '/vendor/autoload.php';

// Being an Iterator, the data can be any dinamically generated content
// for example a PDOStatement set on unbuffered query
$users = new ArrayIterator([
    [
        'column_1' => 'John',
        'column_2' => '123.45',
        'column_3' => '2017-05-08',
    ],
    [
        'column_1' => 'Mary',
        'column_2' => '4321.09',
        'column_3' => '2018-05-08',
    ],
]);

$columnCollection = new ExcelHelper\ColumnCollection([
    new ExcelHelper\Column('column_1',  'User',     10,     new ExcelHelper\CellStyle\Text()),
    new ExcelHelper\Column('column_2',  'Amount',   15,     new ExcelHelper\CellStyle\Amount()),
    new ExcelHelper\Column('column_3',  'Date',     15,     new ExcelHelper\CellStyle\Date()),
]);

$filename = sprintf('%s/my_excel_%s.xls', __DIR__, uniqid());

$phpExcel = new ExcelHelper\TableWorkbook($filename);
$worksheet = $phpExcel->addWorksheet('My Users');

$table = new ExcelHelper\Table($worksheet, 0, 0, 'My Heading', $users);
$table->setColumnCollection($columnCollection);

$phpExcel->writeTable($table);
$phpExcel->close();
```

Result:

![Example](https://raw.githubusercontent.com/Slamdunk/php-excel/master/example.png)
