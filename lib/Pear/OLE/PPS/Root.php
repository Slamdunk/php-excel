<?php

namespace Slam\Excel\Pear\OLE\PPS;

use Slam\Excel;

/**
 * Class for creating Root PPS's for Excel_OLE containers.
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 *
 * @category Structures
 */
class Root extends Excel\Pear\OLE\PPS
{
    protected $_FILEH_;
    protected $_BIG_BLOCK_SIZE;
    protected $_SMALL_BLOCK_SIZE;

    /**
     * Constructor.
     *
     *
     * @param int $time_1st A timestamp
     * @param int $time_2nd A timestamp
     */
    public function __construct($time_1st, $time_2nd, $raChild)
    {
        parent::__construct(
           null,
           Excel\Pear\OLE::Asc2Ucs('Root Entry'),
           Excel\Pear\OLE::Excel_OLE_PPS_TYPE_ROOT,
           null,
           null,
           null,
           $time_1st,
           $time_2nd,
           null,
           $raChild
       );
    }

    /**
     * Method for saving the whole Excel_OLE container (including files).
     * In fact, if called with an empty argument (or '-'), it saves to a
     * temporary file and then outputs it's contents to stdout.
     *
     * @param string $filename The name of the file where to save the Excel_OLE container
     *
     * @return mixed true on success, PEAR_Error on failure
     */
    public function save($filename)
    {
        $this->_FILEH_ = \fopen($filename, 'wb');

        // Initial Setting for saving
        $this->_BIG_BLOCK_SIZE  = \pow(2, ((isset($this->_BIG_BLOCK_SIZE)) ? $this->_adjust2($this->_BIG_BLOCK_SIZE) : 9));
        $this->_SMALL_BLOCK_SIZE = \pow(2, ((isset($this->_SMALL_BLOCK_SIZE)) ? $this->_adjust2($this->_SMALL_BLOCK_SIZE) : 6));

        // Make an array of PPS's (for Save)
        $aList = [];
        Excel\Pear\OLE\PPS\Root::_savePpsSetPnt($aList, [$this]);
        // calculate values for header
        list($iSBDcnt, $iBBcnt, $iPPScnt) = $this->_calcSize($aList); // , $rhInfo);
        // Save Header
        $this->_saveHeader($iSBDcnt, $iBBcnt, $iPPScnt);

        // Make Small Data string (write SBD)
        $this->_data = $this->_makeSmallData($aList);

        // Write BB
        $this->_saveBigData($iSBDcnt, $aList);
        // Write PPS
        $this->_savePps($aList);
        // Write Big Block Depot and BDList and Adding Header informations
        $this->_saveBbd($iSBDcnt, $iBBcnt, $iPPScnt);

        // Close File, send it to stdout if necessary
        \fclose($this->_FILEH_);

        return true;
    }

    /**
     * Calculate some numbers.
     *
     *
     * @param array $raList Reference to an array of PPS's
     *
     * @return array The array of numbers
     */
    private function _calcSize(& $raList)
    {
        // Calculate Basic Setting
        $iSBcnt = 0;
        $iBBcnt = 0;
        foreach ($raList as $pps) {
            if (Excel\Pear\OLE::Excel_OLE_PPS_TYPE_FILE == $pps->Type) {
                $pps->Size = $pps->_DataLen();
                if ($pps->Size < Excel\Pear\OLE::Excel_OLE_DATA_SIZE_SMALL) {
                    $iSBcnt += \floor($pps->Size / $this->_SMALL_BLOCK_SIZE)
                                  + (($pps->Size % $this->_SMALL_BLOCK_SIZE) ? 1 : 0);
                } else {
                    $iBBcnt += (\floor($pps->Size / $this->_BIG_BLOCK_SIZE) +
                        (($pps->Size % $this->_BIG_BLOCK_SIZE) ? 1 : 0));
                }
            }
        }
        $iSmallLen = $iSBcnt * $this->_SMALL_BLOCK_SIZE;
        $iSlCnt = \floor($this->_BIG_BLOCK_SIZE / Excel\Pear\OLE::Excel_OLE_LONG_INT_SIZE);
        $iSBDcnt = \floor($iSBcnt / $iSlCnt) + (($iSBcnt % $iSlCnt) ? 1 : 0);
        $iBBcnt +=  (\floor($iSmallLen / $this->_BIG_BLOCK_SIZE) +
                      (($iSmallLen % $this->_BIG_BLOCK_SIZE) ? 1 : 0));
        $iCnt = \count($raList);
        $iBdCnt = $this->_BIG_BLOCK_SIZE / Excel\Pear\OLE::Excel_OLE_PPS_SIZE;
        $iPPScnt = (\floor($iCnt / $iBdCnt) + (($iCnt % $iBdCnt) ? 1 : 0));

        return [$iSBDcnt, $iBBcnt, $iPPScnt];
    }

    /**
     * Helper function for caculating a magic value for block sizes.
     *
     *
     * @param int $i2 The argument
     *
     * @see save()
     *
     * @return int
     */
    private function _adjust2($i2)
    {
        $iWk = \log($i2) / \log(2);

        return ($iWk > \floor($iWk)) ? \floor($iWk) + 1 : $iWk;
    }

    /**
     * Save Excel_OLE header.
     *
     *
     * @param int $iSBDcnt
     * @param int $iBBcnt
     * @param int $iPPScnt
     */
    private function _saveHeader($iSBDcnt, $iBBcnt, $iPPScnt)
    {
        return $this->_create_header($iSBDcnt, $iBBcnt, $iPPScnt);
    }

    /**
     * Saving big data (PPS's with data bigger than Excel\Pear\OLE::Excel_OLE_DATA_SIZE_SMALL).
     *
     *
     * @param int   $iStBlk
     * @param array $raList Reference to array of PPS's
     */
    private function _saveBigData($iStBlk, & $raList)
    {
        $FILE = $this->_FILEH_;

        // cycle through PPS's
        foreach ($raList as $pps) {
            if (Excel\Pear\OLE::Excel_OLE_PPS_TYPE_DIR == $pps->Type) {
                continue;
            }
            $pps->Size = $pps->_DataLen();
            if (
                    $pps->Size >= Excel\Pear\OLE::Excel_OLE_DATA_SIZE_SMALL
                or  (Excel\Pear\OLE::Excel_OLE_PPS_TYPE_ROOT == $pps->Type and isset($pps->_data))
            ) {
                // Write Data
                if (isset($pps->_PPS_FILE)) {
                    \fseek($pps->_PPS_FILE, 0); // To The Top
                    while ($sBuff = \fread($pps->_PPS_FILE, 4096)) {
                        \fwrite($FILE, $sBuff);
                    }
                } else {
                    \fwrite($FILE, $pps->_data);
                }

                if ($pps->Size % $this->_BIG_BLOCK_SIZE) {
                    for ($j = 0; $j < ($this->_BIG_BLOCK_SIZE - ($pps->Size % $this->_BIG_BLOCK_SIZE)); ++$j) {
                        \fwrite($FILE, "\x00");
                    }
                }
                // Set For PPS
                $pps->_StartBlock = $iStBlk;
                $iStBlk += (
                        \floor($pps->Size / $this->_BIG_BLOCK_SIZE)
                    +   (($pps->Size % $this->_BIG_BLOCK_SIZE) ? 1 : 0)
                );
            }
        }
    }

    /**
     * get small data (PPS's with data smaller than Excel\Pear\OLE::Excel_OLE_DATA_SIZE_SMALL).
     *
     *
     * @param array $raList Reference to array of PPS's
     */
    private function _makeSmallData(& $raList)
    {
        $sRes = '';
        $FILE = $this->_FILEH_;
        $iSmBlk = 0;

        foreach ($raList as $pps) {
            // Make SBD, small data string
            if (
                    Excel\Pear\OLE::Excel_OLE_PPS_TYPE_FILE != $pps->Type
                or  $pps->Size <= 0
                or  $pps->Size >= Excel\Pear\OLE::Excel_OLE_DATA_SIZE_SMALL
            ) {
                continue;
            }

            $iSmbCnt = (
                    \floor($pps->Size / $this->_SMALL_BLOCK_SIZE)
                +   (($pps->Size % $this->_SMALL_BLOCK_SIZE) ? 1 : 0)
            );

            // Add to SBD
            for ($j = 0; $j < ($iSmbCnt - 1); ++$j) {
                \fwrite($FILE, \pack('V', $j + $iSmBlk + 1));
            }
            \fwrite($FILE, \pack('V', -2));

            // Add to Data String(this will be written for RootEntry)
            \fseek($pps->_PPS_FILE, 0); // To The Top
            while ($sBuff = \fread($pps->_PPS_FILE, 4096)) {
                $sRes .= $sBuff;
            }

            if ($pps->Size % $this->_SMALL_BLOCK_SIZE) {
                for ($j = 0; $j < ($this->_SMALL_BLOCK_SIZE - ($pps->Size % $this->_SMALL_BLOCK_SIZE)); ++$j) {
                    $sRes .= "\x00";
                }
            }
            // Set for PPS
            $pps->_StartBlock = $iSmBlk;
            $iSmBlk += $iSmbCnt;
        }

        $iSbCnt = \floor($this->_BIG_BLOCK_SIZE / Excel\Pear\OLE::Excel_OLE_LONG_INT_SIZE);
        if ($iSmBlk % $iSbCnt) {
            for ($i = 0; $i < ($iSbCnt - ($iSmBlk % $iSbCnt)); ++$i) {
                \fwrite($FILE, \pack('V', -1));
            }
        }

        return $sRes;
    }

    /**
     * Saves all the PPS's WKs.
     *
     *
     * @param array $raList Reference to an array with all PPS's
     */
    private function _savePps(& $raList)
    {
        // Save each PPS WK
        foreach ($raList as $pps) {
            \fwrite($this->_FILEH_, $pps->_getPpsWk());
        }
        // Adjust for Block
        $iCnt = \count($raList);
        $iBCnt = $this->_BIG_BLOCK_SIZE / Excel\Pear\OLE::Excel_OLE_PPS_SIZE;
        if ($iCnt % $iBCnt) {
            for ($i = 0; $i < (($iBCnt - ($iCnt % $iBCnt)) * Excel\Pear\OLE::Excel_OLE_PPS_SIZE); ++$i) {
                \fwrite($this->_FILEH_, "\x00");
            }
        }
    }

    /**
     * Saving Big Block Depot.
     *
     *
     * @param int $iSbdSize
     * @param int $iBsize
     * @param int $iPpsCnt
     */
    private function _saveBbd($iSbdSize, $iBsize, $iPpsCnt)
    {
        return $this->_create_big_block_chain($iSbdSize, $iBsize, $iPpsCnt);
    }

    /**
     * New method to store Bigblock chain.
     *
     *
     * @param int $num_sb_blocks  - number of Smallblock depot blocks
     * @param int $num_bb_blocks  - number of Bigblock depot blocks
     * @param int $num_pps_blocks - number of PropertySetStorage blocks
     */
    private function _create_big_block_chain($num_sb_blocks, $num_bb_blocks, $num_pps_blocks)
    {
        $FILE = $this->_FILEH_;

        $bbd_info = $this->_calculate_big_block_chain($num_sb_blocks, $num_bb_blocks, $num_pps_blocks);

        $data = '';

        if ($num_sb_blocks > 0) {
            for ($i = 0; $i < ($num_sb_blocks - 1); ++$i) {
                $data .= \pack('V', $i + 1);
            }
            $data .= \pack('V', -2);
        }

        for ($i = 0; $i < ($num_bb_blocks - 1); ++$i) {
            $data .= \pack('V', $i + $num_sb_blocks + 1);
        }
        $data .= \pack('V', -2);

        for ($i = 0; $i < ($num_pps_blocks - 1); ++$i) {
            $data .= \pack('V', $i + $num_sb_blocks + $num_bb_blocks + 1);
        }
        $data .= \pack('V', -2);

        for ($i = 0; $i < $bbd_info['0xFFFFFFFD_blockchain_entries']; ++$i) {
            $data .= \pack('V', 0xFFFFFFFD);
        }

        for ($i = 0; $i < $bbd_info['0xFFFFFFFC_blockchain_entries']; ++$i) {
            $data .= \pack('V', 0xFFFFFFFC);
        }

        // Adjust for Block
        $all_entries = $num_sb_blocks + $num_bb_blocks + $num_pps_blocks + $bbd_info['0xFFFFFFFD_blockchain_entries'] + $bbd_info['0xFFFFFFFC_blockchain_entries'];
        if ($all_entries % $bbd_info['entries_per_block']) {
            $rest = $bbd_info['entries_per_block'] - ($all_entries % $bbd_info['entries_per_block']);
            for ($i = 0; $i < $rest; ++$i) {
                $data .= \pack('V', -1);
            }
        }

        // Extra BDList
        if ($bbd_info['blockchain_list_entries'] > $bbd_info['header_blockchain_list_entries']) {
            $iN = 0;
            $iNb = 0;
            for ($i = $bbd_info['header_blockchain_list_entries']; $i < $bbd_info['blockchain_list_entries']; $i++, $iN++) {
                if ($iN >= ($bbd_info['entries_per_block'] - 1)) {
                    $iN = 0;
                    ++$iNb;
                    $data .= \pack('V', $num_sb_blocks + $num_bb_blocks + $num_pps_blocks + $bbd_info['0xFFFFFFFD_blockchain_entries'] + $iNb);
                }

                $data .= \pack('V', $num_bb_blocks + $num_sb_blocks + $num_pps_blocks + $i);
            }

            $all_entries = $bbd_info['blockchain_list_entries'] - $bbd_info['header_blockchain_list_entries'];
            if (($all_entries % ($bbd_info['entries_per_block'] - 1))) {
                $rest = ($bbd_info['entries_per_block'] - 1) - ($all_entries % ($bbd_info['entries_per_block'] - 1));
                for ($i = 0; $i < $rest; ++$i) {
                    $data .= \pack('V', -1);
                }
            }

            $data .= \pack('V', -2);
        }

        \fwrite($FILE, $data);
    }

    /**
     * New method to store Header.
     *
     *
     * @param int $num_sb_blocks  - number of Smallblock depot blocks
     * @param int $num_bb_blocks  - number of Bigblock depot blocks
     * @param int $num_pps_blocks - number of PropertySetStorage blocks
     */
    private function _create_header($num_sb_blocks, $num_bb_blocks, $num_pps_blocks)
    {
        $FILE = $this->_FILEH_;

        $bbd_info = $this->_calculate_big_block_chain($num_sb_blocks, $num_bb_blocks, $num_pps_blocks);

        // Save Header
        \fwrite($FILE,
            "\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
            . "\x00\x00\x00\x00"
            . "\x00\x00\x00\x00"
            . "\x00\x00\x00\x00"
            . "\x00\x00\x00\x00"
            . \pack('v', 0x3b)
            . \pack('v', 0x03)
            . \pack('v', -2)
            . \pack('v', 9)
            . \pack('v', 6)
            . \pack('v', 0)
            . "\x00\x00\x00\x00"
            . "\x00\x00\x00\x00"
            . \pack('V', $bbd_info['blockchain_list_entries'])
            . \pack('V', $num_sb_blocks + $num_bb_blocks) // ROOT START
            . \pack('V', 0)
            . \pack('V', 0x1000)
         );

        // Small Block Depot
        if ($num_sb_blocks > 0) {
            \fwrite($FILE, \pack('V', 0));
        } else {
            \fwrite($FILE, \pack('V', -2));
        }

        \fwrite($FILE, \pack('V', $num_sb_blocks));

        // Extra BDList Start, Count
        if ($bbd_info['blockchain_list_entries'] < $bbd_info['header_blockchain_list_entries']) {
            \fwrite($FILE,
                \pack('V', -2) .      // Extra BDList Start
                \pack('V', 0)        // Extra BDList Count
            );
        } else {
            \fwrite($FILE, \pack('V', $num_sb_blocks + $num_bb_blocks + $num_pps_blocks + $bbd_info['0xFFFFFFFD_blockchain_entries']) . \pack('V', $bbd_info['0xFFFFFFFC_blockchain_entries']));
        }

        // BDList
        for ($i = 0; $i < $bbd_info['header_blockchain_list_entries'] and $i < $bbd_info['blockchain_list_entries']; ++$i) {
            \fwrite($FILE, \pack('V', $num_bb_blocks + $num_sb_blocks + $num_pps_blocks + $i));
        }

        if ($i < $bbd_info['header_blockchain_list_entries']) {
            for ($j = 0; $j < ($bbd_info['header_blockchain_list_entries'] - $i); ++$j) {
                \fwrite($FILE, (\pack('V', -1)));
            }
        }
    }

    /**
     * New method to calculate Bigblock chain.
     *
     *
     * @param int $num_sb_blocks  - number of Smallblock depot blocks
     * @param int $num_bb_blocks  - number of Bigblock depot blocks
     * @param int $num_pps_blocks - number of PropertySetStorage blocks
     */
    private function _calculate_big_block_chain($num_sb_blocks, $num_bb_blocks, $num_pps_blocks)
    {
        $bbd_info = [];
        $bbd_info['entries_per_block'] = $this->_BIG_BLOCK_SIZE / Excel\Pear\OLE::Excel_OLE_LONG_INT_SIZE;
        $bbd_info['header_blockchain_list_entries'] = ($this->_BIG_BLOCK_SIZE - 0x4C) / Excel\Pear\OLE::Excel_OLE_LONG_INT_SIZE;
        $bbd_info['blockchain_entries'] = $num_sb_blocks + $num_bb_blocks + $num_pps_blocks;
        $bbd_info['0xFFFFFFFD_blockchain_entries'] = $this->get_number_of_pointer_blocks($bbd_info['blockchain_entries']);
        $bbd_info['blockchain_list_entries'] = $this->get_number_of_pointer_blocks($bbd_info['blockchain_entries'] + $bbd_info['0xFFFFFFFD_blockchain_entries']);

        // do some magic
        $bbd_info['ext_blockchain_list_entries'] = 0;
        $bbd_info['0xFFFFFFFC_blockchain_entries'] = 0;
        if ($bbd_info['blockchain_list_entries'] > $bbd_info['header_blockchain_list_entries']) {
            do {
                $bbd_info['blockchain_list_entries'] = $this->get_number_of_pointer_blocks($bbd_info['blockchain_entries'] + $bbd_info['0xFFFFFFFD_blockchain_entries'] + $bbd_info['0xFFFFFFFC_blockchain_entries']);
                $bbd_info['ext_blockchain_list_entries'] = $bbd_info['blockchain_list_entries'] - $bbd_info['header_blockchain_list_entries'];
                $bbd_info['0xFFFFFFFC_blockchain_entries'] = $this->get_number_of_pointer_blocks($bbd_info['ext_blockchain_list_entries']);
                $bbd_info['0xFFFFFFFD_blockchain_entries'] = $this->get_number_of_pointer_blocks($num_sb_blocks + $num_bb_blocks + $num_pps_blocks + $bbd_info['0xFFFFFFFD_blockchain_entries'] + $bbd_info['0xFFFFFFFC_blockchain_entries']);
            } while ($bbd_info['blockchain_list_entries'] < $this->get_number_of_pointer_blocks($bbd_info['blockchain_entries'] + $bbd_info['0xFFFFFFFD_blockchain_entries'] + $bbd_info['0xFFFFFFFC_blockchain_entries']));
        }

        return $bbd_info;
    }

    /**
     * Calculates number of pointer blocks.
     *
     *
     * @param int $num_pointers - number of pointers
     */
    private function get_number_of_pointer_blocks($num_pointers)
    {
        $pointers_per_block = $this->_BIG_BLOCK_SIZE / Excel\Pear\OLE::Excel_OLE_LONG_INT_SIZE;

        return \floor($num_pointers / $pointers_per_block) + (($num_pointers % $pointers_per_block) ? 1 : 0);
    }
}
