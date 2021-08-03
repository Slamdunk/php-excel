<?php

declare(strict_types=1);

namespace Slam\Excel\Tests\Helper\CellStyle;

use PHPUnit\Framework\TestCase;
use Slam\Excel\Helper\CellStyle\Date;

final class DateTest extends TestCase
{
    /**
     * @dataProvider provide1900Dates
     *
     * @see https://github.com/PHPOffice/PhpSpreadsheet/blob/188d026615f2c79b6a196c97ed8ca82407f47e57/tests/PhpSpreadsheetTests/Shared/DateTest.php#L119-L130
     */
    public function testDateTimeFormattedPHPToExcel1900(?int $expected, ?string $from): void
    {
        self::assertSame($expected, (new Date())->decorateValue($from));
    }

    public function provide1900Dates(): array
    {
        return [
            'null' => [
                null,
                null,
            ],
            'empty-string' => [
                null,
                '',
            ],
            'non-date-string' => [
                null,
                'foobar',
            ],
            'PHP 32-bit Earliest Date 14-Dec-1901' => [
                714,
                '1901-12-14',
            ],
            '31-Dec-1903' => [
                1461,
                '1903-12-31',
            ],
            'Excel 1904 Calendar Base Date   01-Jan-1904' => [
                1462,
                '1904-01-01',
            ],
            '02-Jan-1904' => [
                1463,
                '1904-01-02',
            ],
            '19-Dec-1960' => [
                22269,
                '1960-12-19',
            ],
            'PHP Base Date 01-Jan-1970' => [
                25569,
                '1970-01-01',
            ],
            '07-Dec-1982' => [
                30292,
                '1982-12-07',
            ],
            '12-Jun-2008' => [
                39611,
                '2008-06-12',
            ],
            'PHP 32-bit Latest Date 19-Jan-2038' => [
                50424,
                '2038-01-19',
            ],
        ];
    }
}
