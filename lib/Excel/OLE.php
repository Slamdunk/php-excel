<?php

/**
 * Constants for Excel_OLE package
 */
define('Excel_OLE_PPS_TYPE_ROOT',        5);
define('Excel_OLE_PPS_TYPE_DIR',         1);
define('Excel_OLE_PPS_TYPE_FILE',        2);
define('Excel_OLE_DATA_SIZE_SMALL', 0x1000);
define('Excel_OLE_LONG_INT_SIZE',        4);
define('Excel_OLE_PPS_SIZE',          0x80);

/**
 * Array for storing Excel_OLE instances that are accessed from
 * Excel_OLE_ChainedBlockStream::stream_open().
 *
 * @var  array
 */
$GLOBALS['_Excel_OLE_INSTANCES'] = array();

/**
 * Excel_OLE package base class.
 *
 * @category Structures
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 * @author   Christian Schmidt <schmidt@php.net>
 */
class Excel_OLE
{

    /**
     * The file handle for reading an Excel_OLE container
     *
     * @var resource
     */
    public $_file_handle;

    /**
     * Array of PPS's found on the Excel_OLE container
     *
     * @var array
     */
    public $_list;

    /**
     * Root directory of Excel_OLE container
     *
     * @var Excel_OLE_PPS_Root
     */
    public $root;

    /**
     * Big Block Allocation Table
     *
     * @var array (blockId => nextBlockId)
     */
    public $bbat;

    /**
     * Short Block Allocation Table
     *
     * @var array (blockId => nextBlockId)
     */
    public $sbat;

    /**
     * Size of big blocks. This is usually 512.
     *
     * @var int number of octets per block.
     */
    public $bigBlockSize;

    /**
     * Size of small blocks. This is usually 64.
     *
     * @var int number of octets per block
     */
    public $smallBlockSize;

    // Per unitTests
    public static $gmmktime;
    /**
     * Creates a new Excel_OLE object
     *
     * @access public
     */
    public function Excel_OLE()
    {
        $this->_list = array();
    }

    /**
     * Destructor (using PEAR)
     * Just closes the file handle on the Excel_OLE file.
     *
     * @access private
     */
    public function _Excel_OLE()
    {
        fclose($this->_file_handle);
    }

    /**
     * Reads an Excel_OLE container from the contents of the file given.
     *
     * @access public
     *
     * @param string $file
     *
     * @return mixed true on success, PEAR_Error on failure
     */
    public function read($file)
    {
        $fh = @fopen($file, "r");
        if (! $fh) {
            throw new Excel_Exception_RuntimeException("Can't open file $file");
        }
        $this->_file_handle = $fh;

        $signature = fread($fh, 8);
        if ("\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" != $signature) {
            throw new Excel_Exception_RuntimeException("File doesn't seem to be an Excel_OLE container.");
        }
        fseek($fh, 28);
        if (fread($fh, 2) != "\xFE\xFF") {
            // This shouldn't be a problem in practice
            throw new Excel_Exception_RuntimeException("Only Little-Endian encoding is supported.");
        }
        // Size of blocks and short blocks in bytes
        $this->bigBlockSize   = pow(2, $this->_readInt2($fh));
        $this->smallBlockSize = pow(2, $this->_readInt2($fh));

        // Skip UID, revision number and version number
        fseek($fh, 44);
        // Number of blocks in Big Block Allocation Table
        $bbatBlockCount = $this->_readInt4($fh);

        // Root chain 1st block
        $directoryFirstBlockId = $this->_readInt4($fh);

        // Skip unused bytes
        fseek($fh, 56);
        // Streams shorter than this are stored using small blocks
        $this->bigBlockThreshold = $this->_readInt4($fh);
        // Block id of first sector in Short Block Allocation Table
        $sbatFirstBlockId = $this->_readInt4($fh);
        // Number of blocks in Short Block Allocation Table
        $sbbatBlockCount = $this->_readInt4($fh);
        // Block id of first sector in Master Block Allocation Table
        $mbatFirstBlockId = $this->_readInt4($fh);
        // Number of blocks in Master Block Allocation Table
        $mbbatBlockCount = $this->_readInt4($fh);
        $this->bbat = array();

        // Remaining 4 * 109 bytes of current block is beginning of Master
        // Block Allocation Table
        $mbatBlocks = array();
        for ($i = 0; $i < 109; $i++) {
            $mbatBlocks[] = $this->_readInt4($fh);
        }

        // Read rest of Master Block Allocation Table (if any is left)
        $pos = $this->_getBlockOffset($mbatFirstBlockId);
        for ($i = 0; $i < $mbbatBlockCount; $i++) {
            fseek($fh, $pos);
            for ($j = 0; $j < $this->bigBlockSize / 4 - 1; $j++) {
                $mbatBlocks[] = $this->_readInt4($fh);
            }
            // Last block id in each block points to next block
            $pos = $this->_getBlockOffset($this->_readInt4($fh));
        }

        // Read Big Block Allocation Table according to chain specified by
        // $mbatBlocks
        for ($i = 0; $i < $bbatBlockCount; $i++) {
            $pos = $this->_getBlockOffset($mbatBlocks[$i]);
            fseek($fh, $pos);
            for ($j = 0; $j < $this->bigBlockSize / 4; $j++) {
                $this->bbat[] = $this->_readInt4($fh);
            }
        }

        // Read short block allocation table (SBAT)
        $this->sbat = array();
        $shortBlockCount = $sbbatBlockCount * $this->bigBlockSize / 4;
        $sbatFh = $this->getStream($sbatFirstBlockId);
        if (! $sbatFh) {
            // Avoid an infinite loop if ChainedBlockStream.php somehow is
            // missing
            return false;
        }
        for ($blockId = 0; $blockId < $shortBlockCount; $blockId++) {
            $this->sbat[$blockId] = $this->_readInt4($sbatFh);
        }
        fclose($sbatFh);

        $this->_readPpsWks($directoryFirstBlockId);

        return true;
    }

    /**
     * @param int $blockId block id
     *
     * @return int byte offset from beginning of file
     * @access private
     */
    public function _getBlockOffset($blockId)
    {
        return 512 + $blockId * $this->bigBlockSize;
    }

    /**
     * Returns a stream for use with fread() etc. External callers should
     * use Excel_OLE_PPS_File::getStream().
     *
     * @param int|PPS $blockIdOrPps block id or PPS
     *
     * @return resource read-only stream
     */
    public function getStream($blockIdOrPps)
    {
        include_once 'Excel_OLE/ChainedBlockStream.php';
        static $isRegistered = false;
        if (! $isRegistered) {
            stream_wrapper_register('ole-chainedblockstream',
                                    'Excel_OLE_ChainedBlockStream');
            $isRegistered = true;
        }

        // Store current instance in global array, so that it can be accessed
        // in Excel_OLE_ChainedBlockStream::stream_open().
        // Object is removed from self::$instances in Excel_OLE_Stream::close().
        $GLOBALS['_Excel_OLE_INSTANCES'][] = $this;
        $instanceId = end(array_keys($GLOBALS['_Excel_OLE_INSTANCES']));

        $path = 'ole-chainedblockstream://oleInstanceId=' . $instanceId;
        if (is_a($blockIdOrPps, 'Excel_OLE_PPS')) {
            $path .= '&blockId=' . $blockIdOrPps->_StartBlock;
            $path .= '&size=' . $blockIdOrPps->Size;
        } else {
            $path .= '&blockId=' . $blockIdOrPps;
        }

        return fopen($path, 'r');
    }

    /**
     * Reads a signed char.
     *
     * @param resource $fh file handle
     *
     * @return int
     * @access private
     */
    public function _readInt1($fh)
    {
        list(, $tmp) = unpack("c", fread($fh, 1));

        return $tmp;
    }

    /**
     * Reads an unsigned short (2 octets).
     *
     * @param resource $fh file handle
     *
     * @return int
     * @access private
     */
    public function _readInt2($fh)
    {
        list(, $tmp) = unpack("v", fread($fh, 2));

        return $tmp;
    }

    /**
     * Reads an unsigned long (4 octets).
     *
     * @param   resource  file handle
     *
     * @return int
     * @access private
     */
    public function _readInt4($fh)
    {
        list(, $tmp) = unpack("V", fread($fh, 4));

        return $tmp;
    }

    /**
     * Gets information about all PPS's on the Excel_OLE container from the PPS WK's
     * creates an Excel_OLE_PPS object for each one.
     *
     * @access private
     *
     * @param integer $blockId the block id of the first block
     *
     * @return mixed true on success, PEAR_Error on failure
     */
    public function _readPpsWks($blockId)
    {
        $fh = $this->getStream($blockId);
        for ($pos = 0;; $pos += 128) {
            fseek($fh, $pos, SEEK_SET);
            $nameUtf16 = fread($fh, 64);
            $nameLength = $this->_readInt2($fh);
            $nameUtf16 = substr($nameUtf16, 0, $nameLength - 2);
            // Simple conversion from UTF-16LE to ISO-8859-1
            $name = str_replace("\x00", "", $nameUtf16);
            $type = $this->_readInt1($fh);
            switch ($type) {
            case Excel_OLE_PPS_TYPE_ROOT:
                require_once 'Excel_OLE/PPS/Root.php';
                $pps = new Excel_OLE_PPS_Root(null, null, array());
                $this->root = $pps;
                break;
            case Excel_OLE_PPS_TYPE_DIR:
                $pps = new Excel_OLE_PPS(null, null, null, null, null,
                                   null, null, null, null, array());
                break;
            case Excel_OLE_PPS_TYPE_FILE:
                require_once 'Excel_OLE/PPS/File.php';
                $pps = new Excel_OLE_PPS_File($name);
                break;
            default:
                continue;
            }
            fseek($fh, 1, SEEK_CUR);
            $pps->Type    = $type;
            $pps->Name    = $name;
            $pps->PrevPps = $this->_readInt4($fh);
            $pps->NextPps = $this->_readInt4($fh);
            $pps->DirPps  = $this->_readInt4($fh);
            fseek($fh, 20, SEEK_CUR);
            $pps->Time1st = Excel_OLE::Excel_OLE2LocalDate(fread($fh, 8));
            $pps->Time2nd = Excel_OLE::Excel_OLE2LocalDate(fread($fh, 8));
            $pps->_StartBlock = $this->_readInt4($fh);
            $pps->Size = $this->_readInt4($fh);
            $pps->No = count($this->_list);
            $this->_list[] = $pps;

            // check if the PPS tree (starting from root) is complete
            if (isset($this->root) &&
                $this->_ppsTreeComplete($this->root->No)) {

                break;
            }
        }
        fclose($fh);

        // Initialize $pps->children on directories
        foreach ($this->_list as $pps) {
            if ($pps->Type == Excel_OLE_PPS_TYPE_DIR || $pps->Type == Excel_OLE_PPS_TYPE_ROOT) {
                $nos = array($pps->DirPps);
                $pps->children = array();
                while ($nos) {
                    $no = array_pop($nos);
                    if ($no != -1) {
                        $childPps = $this->_list[$no];
                        $nos[] = $childPps->PrevPps;
                        $nos[] = $childPps->NextPps;
                        $pps->children[] = $childPps;
                    }
                }
            }
        }

        return true;
    }

    /**
     * It checks whether the PPS tree is complete (all PPS's read)
     * starting with the given PPS (not necessarily root)
     *
     * @access private
     *
     * @param integer $index The index of the PPS from which we are checking
     *
     * @return boolean Whether the PPS tree for the given PPS is complete
     */
    public function _ppsTreeComplete($index)
    {
        return isset($this->_list[$index]) &&
               ($pps = $this->_list[$index]) &&
               ($pps->PrevPps == -1 ||
                $this->_ppsTreeComplete($pps->PrevPps)) &&
               ($pps->NextPps == -1 ||
                $this->_ppsTreeComplete($pps->NextPps)) &&
               ($pps->DirPps == -1 ||
                $this->_ppsTreeComplete($pps->DirPps));
    }

    /**
     * Checks whether a PPS is a File PPS or not.
     * If there is no PPS for the index given, it will return false.
     *
     * @param integer $index The index for the PPS
     *
     * @return bool true if it's a File PPS, false otherwise
     * @access public
     */
    public function isFile($index)
    {
        if (isset($this->_list[$index])) {
            return ($this->_list[$index]->Type == Excel_OLE_PPS_TYPE_FILE);
        }

        return false;
    }

    /**
     * Checks whether a PPS is a Root PPS or not.
     * If there is no PPS for the index given, it will return false.
     *
     * @param integer $index The index for the PPS.
     *
     * @return bool true if it's a Root PPS, false otherwise
     * @access public
     */
    public function isRoot($index)
    {
        if (isset($this->_list[$index])) {
            return ($this->_list[$index]->Type == Excel_OLE_PPS_TYPE_ROOT);
        }

        return false;
    }

    /**
     * Gives the total number of PPS's found in the Excel_OLE container.
     *
     * @return integer The total number of PPS's found in the Excel_OLE container
     * @access public
     */
    public function ppsTotal()
    {
        return count($this->_list);
    }

    /**
     * Gets data from a PPS
     * If there is no PPS for the index given, it will return an empty string.
     *
     * @param integer $index    The index for the PPS
     * @param integer $position The position from which to start reading
     *                          (relative to the PPS)
     * @param integer $length   The amount of bytes to read (at most)
     *
     * @return string The binary string containing the data requested
     * @access public
     *
     * @see Excel_OLE_PPS_File::getStream()
     */
    public function getData($index, $position, $length)
    {
        // if position is not valid return empty string
        if (! isset($this->_list[$index]) ||
            $position >= $this->_list[$index]->Size ||
            $position < 0) {

            return '';
        }
        $fh = $this->getStream($this->_list[$index]);
        $data = stream_get_contents($fh, $length, $position);
        fclose($fh);

        return $data;
    }

    /**
     * Gets the data length from a PPS
     * If there is no PPS for the index given, it will return 0.
     *
     * @param integer $index The index for the PPS
     *
     * @return integer The amount of bytes in data the PPS has
     * @access public
     */
    public function getDataLength($index)
    {
        if (isset($this->_list[$index])) {
            return $this->_list[$index]->Size;
        }

        return 0;
    }

    /**
     * Utility function to transform ASCII text to Unicode
     *
     * @access public
     * @static
     *
     * @param string $ascii The ASCII string to transform
     *
     * @return string The string in Unicode
     */
    public static function Asc2Ucs($ascii)
    {
        $rawname = '';
        for ($i = 0; $i < strlen($ascii); $i++) {
            $rawname .= $ascii{$i} . "\x00";
        }

        return $rawname;
    }

    /**
     * Utility function
     * Returns a string for the Excel_OLE container with the date given
     *
     * @access public
     * @static
     *
     * @param integer $date A timestamp
     *
     * @return string The string for the Excel_OLE container
     */
    public static function LocalDate2Excel_OLE($date = null)
    {
        if (! isset($date)) {
            return "\x00\x00\x00\x00\x00\x00\x00\x00";
        }

        // factor used for separating numbers into 4 bytes parts
        $factor = pow(2, 32);

        // days from 1-1-1601 until the beggining of UNIX era
        $days = 134774;
        // calculate seconds
        $gmmktime = gmmktime(
            date("H",$date),date("i",$date),date("s",$date),
            date("m",$date),date("d",$date),date("Y",$date)
        );
        if (isset(self::$gmmktime)) {
            $gmmktime = self::$gmmktime;
        }
        $big_date = $days * 24 * 3600 + $gmmktime;
        // multiply just to make MS happy
        $big_date *= 10000000;

        $high_part = floor($big_date / $factor);
        // lower 4 bytes
        $low_part = floor((($big_date / $factor) - $high_part) * $factor);

        // Make HEX string
        $res = '';

        for ($i = 0; $i < 4; $i++) {
            $hex = $low_part % 0x100;
            $res .= pack('c', $hex);
            $low_part /= 0x100;
        }
        for ($i = 0; $i < 4; $i++) {
            $hex = $high_part % 0x100;
            $res .= pack('c', $hex);
            $high_part /= 0x100;
        }

        return $res;
    }

    /**
     * Returns a timestamp from an Excel_OLE container's date
     *
     * @param integer $string A binary string with the encoded date
     *
     * @return string The timestamp corresponding to the string
     * @access public
     * @static
     */
    public static function Excel_OLE2LocalDate($string)
    {
        if (strlen($string) != 8) {
            return new Excel_PEAR_Error("Expecting 8 byte string");
        }

        // factor used for separating numbers into 4 bytes parts
        $factor = pow(2,32);
        $high_part = 0;
        for ($i = 0; $i < 4; $i++) {
            list(, $high_part) = unpack('C', $string{(7 - $i)});
            if ($i < 3) {
                $high_part *= 0x100;
            }
        }
        $low_part = 0;
        for ($i = 4; $i < 8; $i++) {
            list(, $low_part) = unpack('C', $string{(7 - $i)});
            if ($i < 7) {
                $low_part *= 0x100;
            }
        }
        $big_date = ($high_part * $factor) + $low_part;
        // translate to seconds
        $big_date /= 10000000;

        // days from 1-1-1601 until the beggining of UNIX era
        $days = 134774;

        // translate to seconds from beggining of UNIX era
        $big_date -= $days * 24 * 3600;

        return floor($big_date);
    }
}
