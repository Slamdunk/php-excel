<?php

/**
* Class for creating PPS's for Excel_OLE containers
*
* @author   Xavier Noguer <xnoguer@php.net>
* @category Structures
* @package  Excel_OLE
*/
class Excel_OLE_PPS
{
    /**
    * The PPS index
    * @var integer
    */
    var $No;

    /**
    * The PPS name (in Unicode)
    * @var string
    */
    var $Name;
 
    /**
    * The PPS type. Dir, Root or File
    * @var integer
    */
    var $Type;
 
    /**
    * The index of the previous PPS
    * @var integer
    */
    var $PrevPps;
 
    /**
    * The index of the next PPS
    * @var integer
    */
    var $NextPps;
 
    /**
    * The index of it's first child if this is a Dir or Root PPS
    * @var integer
    */
    var $DirPps;
 
    /**
    * A timestamp
    * @var integer
    */
    var $Time1st;

    /**
    * A timestamp
    * @var integer
    */
    var $Time2nd;

    /**
    * Starting block (small or big) for this PPS's data  inside the container
    * @var integer
    */
    var $_StartBlock;

    /**
    * The size of the PPS's data (in bytes)
    * @var integer
    */
    var $Size;

    /**
    * The PPS's data (only used if it's not using a temporary file)
    * @var string
    */
    var $_data;

    /**
    * Array of child PPS's (only used by Root and Dir PPS's)
    * @var array
    */
    var $children = array();

    /**
    * Pointer to Excel_OLE container
    * @var Excel_OLE
    */
    var $ole;

    /**
    * The constructor
    *
    * @access public
    * @param integer $No   The PPS index
    * @param string  $name The PPS name
    * @param integer $type The PPS type. Dir, Root or File
    * @param integer $prev The index of the previous PPS
    * @param integer $next The index of the next PPS
    * @param integer $dir  The index of it's first child if this is a Dir or Root PPS
    * @param integer $time_1st A timestamp
    * @param integer $time_2nd A timestamp
    * @param string  $data  The (usually binary) source data of the PPS
    * @param array   $children Array containing children PPS for this PPS
    */
    function Excel_OLE_PPS($No, $name, $type, $prev, $next, $dir, $time_1st, $time_2nd, $data, $children)
    {
        $this->No      = $No;
        $this->Name    = $name;
        $this->Type    = $type;
        $this->PrevPps = $prev;
        $this->NextPps = $next;
        $this->DirPps  = $dir;
        $this->Time1st = $time_1st;
        $this->Time2nd = $time_2nd;
        $this->_data      = $data;
        $this->children   = $children;
        if ($data != '') {
            $this->Size = strlen($data);
        } else {
            $this->Size = 0;
        }
    }

    /**
    * Returns the amount of data saved for this PPS
    *
    * @access private
    * @return integer The amount of data (in bytes)
    */
    function _DataLen()
    {
        if (!isset($this->_data)) {
            return 0;
        }
        if (isset($this->_PPS_FILE)) {
            fseek($this->_PPS_FILE, 0);
            $stats = fstat($this->_PPS_FILE);
            return $stats[7];
        } else {
            return strlen($this->_data);
        }
    }

    /**
    * Returns a string with the PPS's WK (What is a WK?)
    *
    * @access private
    * @return string The binary string
    */
    function _getPpsWk()
    {
        $ret = $this->Name;
        for ($i = 0; $i < (64 - strlen($this->Name)); $i++) {
            $ret .= "\x00";
        }
        $ret .= pack("v", strlen($this->Name) + 2)  // 66
              . pack("c", $this->Type)              // 67
              . pack("c", 0x00) //UK                // 68
              . pack("V", $this->PrevPps) //Prev    // 72
              . pack("V", $this->NextPps) //Next    // 76
              . pack("V", $this->DirPps)  //Dir     // 80
              . "\x00\x09\x02\x00"                  // 84
              . "\x00\x00\x00\x00"                  // 88
              . "\xc0\x00\x00\x00"                  // 92
              . "\x00\x00\x00\x46"                  // 96 // Seems to be ok only for Root
              . "\x00\x00\x00\x00"                  // 100
              . Excel_OLE::LocalDate2Excel_OLE($this->Time1st)       // 108
              . Excel_OLE::LocalDate2Excel_OLE($this->Time2nd)       // 116
              . pack("V", isset($this->_StartBlock)? 
                        $this->_StartBlock:0)        // 120
              . pack("V", $this->Size)               // 124
              . pack("V", 0);                        // 128
        return $ret;
    }

    /**
    * Updates index and pointers to previous, next and children PPS's for this
    * PPS. I don't think it'll work with Dir PPS's.
    *
    * @access private
    * @param array &$pps_array Reference to the array of PPS's for the whole Excel_OLE
    *                          container 
    * @return integer          The index for this PPS
    */
    static function _savePpsSetPnt(&$raList, $to_save, $depth = 0) 
    {
      if ( !is_array($to_save) || (count($to_save) == 0) ) {
        return 0xFFFFFFFF;
      }
      elseif( count($to_save) == 1 ) {
        $cnt = count($raList);
        // If the first entry, it's the root... Don't clone it!
        $raList[$cnt] = ( $depth == 0 ) ? $to_save[0] : clone $to_save[0];
        $raList[$cnt]->No = $cnt;
        $raList[$cnt]->PrevPps = 0xFFFFFFFF;
        $raList[$cnt]->NextPps = 0xFFFFFFFF;
        $raList[$cnt]->DirPps  = self::_savePpsSetPnt($raList, @$raList[$cnt]->children, $depth++);
        return $cnt;
      }
      else {
        $iPos  = floor(count($to_save) / 2);
        $aPrev = array_slice($to_save, 0, $iPos);
        $aNext = array_slice($to_save, $iPos + 1);

        $cnt   = count($raList);
        // If the first entry, it's the root... Don't clone it!
        $raList[$cnt] = ( $depth == 0 ) ? $to_save[$iPos] : clone $to_save[$iPos];
        $raList[$cnt]->No = $cnt;
        $raList[$cnt]->PrevPps = self::_savePpsSetPnt($raList, $aPrev, $depth++);
        $raList[$cnt]->NextPps = self::_savePpsSetPnt($raList, $aNext, $depth++);
        $raList[$cnt]->DirPps  = self::_savePpsSetPnt($raList, @$raList[$cnt]->children, $depth++);

        return $cnt;
      }
    }
}
?>
