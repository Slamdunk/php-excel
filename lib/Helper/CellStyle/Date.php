<?php

declare(strict_types=1);

namespace Slam\Excel\Helper\CellStyle;

use Slam\Excel\Helper\CellStyleInterface;
use Slam\Excel\Pear\Writer\Format;

final class Date implements CellStyleInterface
{
    public function decorateValue($value): ?int
    {
        if (! \is_string($value) || 1 !== \preg_match('/^(?<year>\d\d\d\d)-(?<month>\d\d)-(?<day>\d\d)$/', $value, $matches)) {
            return null;
        }

        $year  = (int) $matches['year'];
        $month = (int) $matches['month'];
        $day   = (int) $matches['day'];

        //    Fudge factor for the erroneous fact that the year 1900 is treated as a Leap Year in MS Excel
        //    This affects every date following 28th February 1900
        $excel1900isLeapYear = true;
        if ((1900 === $year) && ($month <= 2)) {
            $excel1900isLeapYear = false;
        }
        $myexcelBaseDate = 2415020;

        //    Julian base date Adjustment
        if ($month > 2) {
            $month -= 3;
        } else {
            $month += 9;
            --$year;
        }

        //    Calculate the Julian Date, then subtract the Excel base date (JD 2415020 = 31-Dec-1899 Giving Excel Date of 0)
        $century   = (int) \substr((string) $year, 0, 2);
        $decade    = (int) \substr((string) $year, 2, 2);
        $excelDate = \floor((146097 * $century) / 4) + \floor((1461 * $decade) / 4) + \floor((153 * $month + 2) / 5) + $day + 1721119 - $myexcelBaseDate + $excel1900isLeapYear;

        return (int) $excelDate;
    }

    public function styleCell(Format $format): void
    {
        $format->setAlign('center');
        $format->setNumFormat('DD/MM/YYYY');
    }
}
