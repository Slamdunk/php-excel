<?php

namespace Slam\Excel\Pear\Writer;

use Slam\Excel;

/**
 * Class for writing Excel BIFF records.
 *
 * From "MICROSOFT EXCEL BINARY FILE FORMAT" by Mark O'Brien (Microsoft Corporation):
 *
 * BIFF (BInary File Format) is the file format in which Excel documents are
 * saved on disk.  A BIFF file is a complete description of an Excel document.
 * BIFF files consist of sequences of variable-length records. There are many
 * different types of BIFF records.  For example, one record type describes a
 * formula entered into a cell; one describes the size and location of a
 * window into a document; another describes a picture format.
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 *
 * @category FileFormats
 */
class BIFFwriter
{
    /**
     * The BIFF/Excel version (5).
     *
     * @var int
     */
    const BIFF_version = 0x0500;

    /**
     * The byte order of this architecture. 0 => little endian, 1 => big endian.
     *
     * @var int
     */
    protected $_byte_order;

    /**
     * The string containing the data of the BIFF stream.
     *
     * @var string
     */
    protected $_data;

    /**
     * The size of the data in bytes. Should be the same as strlen($this->_data).
     *
     * @var int
     */
    protected $_datasize;

    /**
     * The maximun length for a BIFF record. See _addContinue().
     *
     * @var int
     *
     * @see _addContinue()
     */
    protected $_limit;

    /**
     * Constructor.
     */
    public function __construct()
    {
        $this->_data       = '';
        $this->_datasize   = 0;
        $this->_limit      = 2080;
        // Set the byte order
        $this->_setByteOrder();
    }

    /**
     * Determine the byte order and store it as class data to avoid
     * recalculating it for each call to new().
     */
    protected function _setByteOrder()
    {
        // Check if "pack" gives the required IEEE 64bit float
        $teststr = \pack('d', 1.2345);
        $number  = \pack('C8', 0x8D, 0x97, 0x6E, 0x12, 0x83, 0xC0, 0xF3, 0x3F);
        if ($number == $teststr) {
            $byte_order = 0;    // Little Endian
        } elseif ($number == \strrev($teststr)) {
            $byte_order = 1;    // Big Endian
        } else {
            // Give up. I'll fix this in a later version.
            throw new Excel\Exception\RuntimeException('Required floating point format not supported on this platform.');
        }
        $this->_byte_order = $byte_order;
    }

    /**
     * General storage function.
     *
     * @param string $data binary data to prepend
     */
    protected function _prepend($data)
    {
        if (\strlen($data) > $this->_limit) {
            $data = $this->_addContinue($data);
        }
        $this->_data      = $data . $this->_data;
        $this->_datasize += \strlen($data);
    }

    /**
     * General storage function.
     *
     * @param string $data binary data to append
     */
    protected function _append($data)
    {
        if (\strlen($data) > $this->_limit) {
            $data = $this->_addContinue($data);
        }
        $this->_data      = $this->_data . $data;
        $this->_datasize += \strlen($data);
    }

    /**
     * Writes Excel BOF record to indicate the beginning of a stream or
     * sub-stream in the BIFF file.
     *
     * @param int $type type of BIFF file to write: 0x0005 Workbook,
     *                  0x0010 Worksheet
     */
    protected function _storeBof($type)
    {
        $record  = 0x0809;        // Record identifier

        // According to the SDK $build and $year should be set to zero.
        // However, this throws a warning in Excel 5. So, use magic numbers.
        $length  = 0x0008;
        $unknown = '';
        $build   = 0x096C;
        $year    = 0x07C9;

        $version = self::BIFF_version;

        $header  = \pack('vv',   $record, $length);
        $data    = \pack('vvvv', $version, $type, $build, $year);
        $this->_prepend($header . $data . $unknown);
    }

    /**
     * Writes Excel EOF record to indicate the end of a BIFF stream.
     */
    protected function _storeEof()
    {
        $record    = 0x000A;   // Record identifier
        $length    = 0x0000;   // Number of bytes to follow
        $header    = \pack('vv', $record, $length);
        $this->_append($header);
    }

    /**
     * Excel limits the size of BIFF records. In Excel 5 the limit is 2084 bytes. In
     * Excel 97 the limit is 8228 bytes. Records that are longer than these limits
     * must be split up into CONTINUE blocks.
     *
     * This function takes a long BIFF record and inserts CONTINUE records as
     * necessary.
     *
     * @param string $data The original binary data to be written
     *
     * @return string A very convenient string of continue blocks
     */
    protected function _addContinue($data)
    {
        $limit  = $this->_limit;
        $record = 0x003C;         // Record identifier

        // The first 2080/8224 bytes remain intact. However, we have to change
        // the length field of the record.
        $tmp = \substr($data, 0, 2) . \pack('v', $limit - 4) . \substr($data, 4, $limit - 4);

        $header = \pack('vv', $record, $limit);  // Headers for continue records

        // Retrieve chunks of 2080/8224 bytes +4 for the header.
        $data_length = \strlen($data);
        for ($i = $limit; $i <  ($data_length - $limit); $i += $limit) {
            $tmp .= $header;
            $tmp .= \substr($data, $i, $limit);
        }

        // Retrieve the last chunk of data
        $header  = \pack('vv', $record, \strlen($data) - $i);
        $tmp    .= $header;
        $tmp    .= \substr($data, $i, \strlen($data) - $i);

        return $tmp;
    }
}
