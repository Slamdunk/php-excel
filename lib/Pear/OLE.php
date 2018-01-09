<?php

namespace Slam\Excel\Pear;

use Slam\Excel\Exception;

/**
 * Excel_OLE package base class.
 *
 * @category Structures
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 * @author   Christian Schmidt <schmidt@php.net>
 */
class OLE
{
    const Excel_OLE_DATA_SIZE_SMALL     = 0x1000;
    const Excel_OLE_LONG_INT_SIZE       = 4;
    const Excel_OLE_PPS_SIZE            = 0x80;
    const Excel_OLE_PPS_TYPE_DIR        = 1;
    const Excel_OLE_PPS_TYPE_FILE       = 2;
    const Excel_OLE_PPS_TYPE_ROOT       = 5;

    // For Unit Tests
    public static $gmmktime;

    public static function getTmpfile()
    {
        $resource = \tmpfile();
        if (! \is_resource($resource)) {
            throw new Exception\RuntimeException('Can\'t create temporary file');
        }

        return $resource;
    }

    /**
     * Utility function to transform ASCII text to Unicode.
     *
     * @static
     *
     * @param string $ascii The ASCII string to transform
     *
     * @return string The string in Unicode
     */
    public static function Asc2Ucs($ascii)
    {
        $rawname = '';
        for ($i = 0; $i < \strlen($ascii); ++$i) {
            $rawname .= $ascii[$i] . "\x00";
        }

        return $rawname;
    }

    /**
     * Utility function
     * Returns a string for the Excel_OLE container with the date given.
     *
     * @static
     *
     * @param int $date A timestamp
     *
     * @return string The string for the Excel_OLE container
     */
    public static function LocalDate2Excel_OLE($date = null)
    {
        if (! isset($date)) {
            return "\x00\x00\x00\x00\x00\x00\x00\x00";
        }

        // factor used for separating numbers into 4 bytes parts
        $factor = \pow(2, 32);

        // days from 1-1-1601 until the beggining of UNIX era
        $days = 134774;
        // calculate seconds
        $gmmktime = \gmmktime(
            \date('H', $date), \date('i', $date), \date('s', $date),
            \date('m', $date), \date('d', $date), \date('Y', $date)
        );
        if (isset(self::$gmmktime)) {
            $gmmktime = self::$gmmktime;
        }
        $big_date = $days * 24 * 3600 + $gmmktime;
        // multiply just to make MS happy
        $big_date *= 10000000;

        $high_part = \floor($big_date / $factor);
        // lower 4 bytes
        $low_part = \floor((($big_date / $factor) - $high_part) * $factor);

        // Make HEX string
        $res = '';

        for ($i = 0; $i < 4; ++$i) {
            $hex = $low_part % 0x100;
            $res .= \pack('c', $hex);
            $low_part /= 0x100;
        }
        for ($i = 0; $i < 4; ++$i) {
            $hex = $high_part % 0x100;
            $res .= \pack('c', $hex);
            $high_part /= 0x100;
        }

        return $res;
    }
}
