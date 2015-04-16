<?php

/**
 * Class for writing Excel Spreadsheets. This class should change COMPLETELY.
 *
 * @author   Xavier Noguer <xnoguer@rezebra.com>
 *
 * @category FileFormats
 */

class Excel_Writer extends Excel_Writer_Workbook
{
    /**
     * The constructor. It just creates a Workbook
     *
     * @param string $filename The optional filename for the Workbook.
     *
     * @return Excel_Writer_Workbook The Workbook created
     */
    public function Excel_Writer($filename = '')
    {
        $this->_filename = $filename;
        $this->Excel_Writer_Workbook($filename);
    }

    /**
     * Utility function for writing formulas
     * Converts a cell's coordinates to the A1 format.
     *
     * @access public
     * @static
     *
     * @param integer $row Row for the cell to convert (0-indexed).
     * @param integer $col Column for the cell to convert (0-indexed).
     *
     * @return string The cell identifier in A1 format
     */
    public function rowcolToCell($row, $col)
    {
        if ($col > 255) { //maximum column value exceeded
            throw new Excel_Exception_InvalidArgumentException("Maximum column value exceeded: $col");
        }

        $int = (int) ($col / 26);
        $frac = $col % 26;
        $chr1 = '';

        if ($int > 0) {
            $chr1 = chr(ord('A') + $int - 1);
        }

        $chr2 = chr(ord('A') + $frac);
        $row++;

        return $chr1 . $chr2 . $row;
    }
}
