<?php

/**
* Class for generating Excel XF records (formats)
*
* @author   Xavier Noguer <xnoguer@rezebra.com>
* @category FileFormats
* @package  Excel_Writer
*/

class Excel_Writer_Format extends Excel_PEAR
{
    /**
    * The index given by the workbook when creating a new format.
    * @var integer
    */
    var $_xf_index;

    /**
    * Index to the FONT record.
    * @var integer
    */
    var $font_index;

    /**
    * The font name (ASCII).
    * @var string
    */
    var $_font_name;

    /**
    * Height of font (1/20 of a point)
    * @var integer
    */
    var $_size;

    /**
    * Bold style
    * @var integer
    */
    var $_bold;

    /**
    * Bit specifiying if the font is italic.
    * @var integer
    */
    var $_italic;

    /**
    * Index to the cell's color
    * @var integer
    */
    var $_color;

    /**
    * The text underline property
    * @var integer
    */
    var $_underline;

    /**
    * Bit specifiying if the font has strikeout.
    * @var integer
    */
    var $_font_strikeout;

    /**
    * Bit specifiying if the font has outline.
    * @var integer
    */
    var $_font_outline;

    /**
    * Bit specifiying if the font has shadow.
    * @var integer
    */
    var $_font_shadow;

    /**
    * 2 bytes specifiying the script type for the font.
    * @var integer
    */
    var $_font_script;

    /**
    * Byte specifiying the font family.
    * @var integer
    */
    var $_font_family;

    /**
    * Byte specifiying the font charset.
    * @var integer
    */
    var $_font_charset;

    /**
    * An index (2 bytes) to a FORMAT record (number format).
    * @var integer
    */
    var $_num_format;

    /**
    * Bit specifying if formulas are hidden.
    * @var integer
    */
    var $_hidden;

    /**
    * Bit specifying if the cell is locked.
    * @var integer
    */
    var $_locked;

    /**
    * The three bits specifying the text horizontal alignment.
    * @var integer
    */
    var $_text_h_align;

    /**
    * Bit specifying if the text is wrapped at the right border.
    * @var integer
    */
    var $_text_wrap;

    /**
    * The three bits specifying the text vertical alignment.
    * @var integer
    */
    var $_text_v_align;

    /**
    * 1 bit, apparently not used.
    * @var integer
    */
    var $_text_justlast;

    /**
    * The two bits specifying the text rotation.
    * @var integer
    */
    var $_rotation;

    /**
    * The cell's foreground color.
    * @var integer
    */
    var $_fg_color;

    /**
    * The cell's background color.
    * @var integer
    */
    var $_bg_color;

    /**
    * The cell's background fill pattern.
    * @var integer
    */
    var $_pattern;

    /**
    * Style of the bottom border of the cell
    * @var integer
    */
    var $_bottom;

    /**
    * Color of the bottom border of the cell.
    * @var integer
    */
    var $_bottom_color;

    /**
    * Style of the top border of the cell
    * @var integer
    */
    var $_top;

    /**
    * Color of the top border of the cell.
    * @var integer
    */
    var $_top_color;

    /**
    * Style of the left border of the cell
    * @var integer
    */
    var $_left;

    /**
    * Color of the left border of the cell.
    * @var integer
    */
    var $_left_color;

    /**
    * Style of the right border of the cell
    * @var integer
    */
    var $_right;

    /**
    * Color of the right border of the cell.
    * @var integer
    */
    var $_right_color;

    /**
    * Constructor
    *
    * @access private
    * @param integer $index the XF index for the format.
    * @param array   $properties array with properties to be set on initialization.
    */
    function Excel_Writer_Format($index = 0, $properties =  array())
    {
        $this->_xf_index       = $index;
        $this->font_index      = 0;
        $this->_font_name      = 'Arial';
        $this->_size           = 10;
        $this->_bold           = 0x0190;
        $this->_italic         = 0;
        $this->_color          = 0x7FFF;
        $this->_underline      = 0;
        $this->_font_strikeout = 0;
        $this->_font_outline   = 0;
        $this->_font_shadow    = 0;
        $this->_font_script    = 0;
        $this->_font_family    = 0;
        $this->_font_charset   = 0;

        $this->_num_format     = 0;

        $this->_hidden         = 0;
        $this->_locked         = 0;

        $this->_text_h_align   = 0;
        $this->_text_wrap      = 0;
        $this->_text_v_align   = 2;
        $this->_text_justlast  = 0;
        $this->_rotation       = 0;

        $this->_fg_color       = 0x40;
        $this->_bg_color       = 0x41;

        $this->_pattern        = 0;

        $this->_bottom         = 0;
        $this->_top            = 0;
        $this->_left           = 0;
        $this->_right          = 0;
        $this->_diag           = 0;

        $this->_bottom_color   = 0x40;
        $this->_top_color      = 0x40;
        $this->_left_color     = 0x40;
        $this->_right_color    = 0x40;
        $this->_diag_color     = 0x40;

        // Set properties passed to Excel_Writer_Workbook::addFormat()
        foreach ($properties as $property => $value)
        {
            if (method_exists($this, 'set'.ucwords($property))) {
                $method_name = 'set'.ucwords($property);
                $this->$method_name($value);
            }
        }
    }


    /**
    * Generate an Excel BIFF XF record (style or cell).
    *
    * @param string $style The type of the XF record ('style' or 'cell').
    * @return string The XF record
    */
    function getXf($style)
    {
        // Set the type of the XF record and some of the attributes.
        if ($style == 'style') {
            $style = 0xFFF5;
        } else {
            $style   = $this->_locked;
            $style  |= $this->_hidden << 1;
        }

        // Flags to indicate if attributes have been set.
        $atr_num     = ($this->_num_format != 0)?1:0;
        $atr_fnt     = ($this->font_index != 0)?1:0;
        $atr_alc     = ($this->_text_wrap)?1:0;
        $atr_bdr     = ($this->_bottom   ||
                        $this->_top      ||
                        $this->_left     ||
                        $this->_right)?1:0;
        $atr_pat     = (($this->_fg_color != 0x40) ||
                        ($this->_bg_color != 0x41) ||
                        $this->_pattern)?1:0;
        $atr_prot    = $this->_locked | $this->_hidden;

        // Zero the default border colour if the border has not been set.
        if ($this->_bottom == 0) {
            $this->_bottom_color = 0;
        }
        if ($this->_top  == 0) {
            $this->_top_color = 0;
        }
        if ($this->_right == 0) {
            $this->_right_color = 0;
        }
        if ($this->_left == 0) {
            $this->_left_color = 0;
        }
        if ($this->_diag == 0) {
            $this->_diag_color = 0;
        }

        $record         = 0x00E0;              // Record identifier
        $length         = 0x0010;              // Number of bytes to follow
        
        $ifnt           = $this->font_index;   // Index to FONT record
        $ifmt           = $this->_num_format;  // Index to FORMAT record

        $align          = $this->_text_h_align;       // Alignment
        $align         |= $this->_text_wrap     << 3;
        $align         |= $this->_text_v_align  << 4;
        $align         |= $this->_text_justlast << 7;
        $align         |= $this->_rotation      << 8;
        $align         |= $atr_num                << 10;
        $align         |= $atr_fnt                << 11;
        $align         |= $atr_alc                << 12;
        $align         |= $atr_bdr                << 13;
        $align         |= $atr_pat                << 14;
        $align         |= $atr_prot               << 15;

        $icv            = $this->_fg_color;       // fg and bg pattern colors
        $icv           |= $this->_bg_color      << 7;

        $fill           = $this->_pattern;        // Fill and border line style
        $fill          |= $this->_bottom        << 6;
        $fill          |= $this->_bottom_color  << 9;

        $border1        = $this->_top;            // Border line style and color
        $border1       |= $this->_left          << 3;
        $border1       |= $this->_right         << 6;
        $border1       |= $this->_top_color     << 9;

        $border2        = $this->_left_color;     // Border color
        $border2       |= $this->_right_color   << 7;

        $header      = pack("vv",       $record, $length);
        $data        = pack("vvvvvvvv", $ifnt, $ifmt, $style, $align,
                                        $icv, $fill,
                                        $border1, $border2);

        return($header . $data);
    }

    /**
    * Generate an Excel BIFF FONT record.
    *
    * @return string The FONT record
    */
    function getFont()
    {
        $dyHeight   = $this->_size * 20;    // Height of font (1/20 of a point)
        $icv        = $this->_color;        // Index to color palette
        $bls        = $this->_bold;         // Bold style
        $sss        = $this->_font_script;  // Superscript/subscript
        $uls        = $this->_underline;    // Underline
        $bFamily    = $this->_font_family;  // Font family
        $bCharSet   = $this->_font_charset; // Character set
        $encoding   = 0;                    // TODO: Unicode support

        $cch        = strlen($this->_font_name); // Length of font name
        $record     = 0x31;                      // Record identifier
        $length     = 0x0F + $cch;            // Record length
        $reserved   = 0x00;                // Reserved
        $grbit      = 0x00;                // Font attributes
        if ($this->_italic) {
            $grbit     |= 0x02;
        }
        if ($this->_font_strikeout) {
            $grbit     |= 0x08;
        }
        if ($this->_font_outline) {
            $grbit     |= 0x10;
        }
        if ($this->_font_shadow) {
            $grbit     |= 0x20;
        }

        $header  = pack("vv",         $record, $length);
        $data    = pack("vvvvvCCCCC", $dyHeight, $grbit, $icv, $bls,
                                      $sss, $uls, $bFamily,
                                      $bCharSet, $reserved, $cch);

        return($header . $data . $this->_font_name);
    }

    /**
    * Returns a unique hash key for a font.
    * Used by Excel_Writer_Workbook::_storeAllFonts()
    *
    * The elements that form the key are arranged to increase the probability of
    * generating a unique key. Elements that hold a large range of numbers
    * (eg. _color) are placed between two binary elements such as _italic
    *
    * @return string A key for this font
    */
    function getFontKey()
    {
        $key  = "$this->_font_name$this->_size";
        $key .= "$this->_font_script$this->_underline";
        $key .= "$this->_font_strikeout$this->_bold$this->_font_outline";
        $key .= "$this->_font_family$this->_font_charset";
        $key .= "$this->_font_shadow$this->_color$this->_italic";
        $key  = str_replace(' ', '_', $key);
        return ($key);
    }

    /**
    * Returns the index used by Excel_Writer_Worksheet::_XF()
    *
    * @return integer The index for the XF record
    */
    function getXfIndex()
    {
        return($this->_xf_index);
    }

    /**
    * Used in conjunction with the set_xxx_color methods to convert a color
    * string into a number. Color range is 0..63 but we will restrict it
    * to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
    *
    * @access private
    * @param string $name_color name of the color (i.e.: 'blue', 'red', etc..). Optional.
    * @return integer The color index
    */
    function _getColor($name_color = '')
    {
        $colors = array(
                          'aqua'    => 0x07,
                          'cyan'    => 0x07,
                          'black'   => 0x00,
                          'blue'    => 0x04,
                          'brown'   => 0x10,
                          'magenta' => 0x06,
                          'fuchsia' => 0x06,
                          'gray'    => 0x17,
                          'grey'    => 0x17,
                          'green'   => 0x11,
                          'lime'    => 0x03,
                          'navy'    => 0x12,
                          'orange'  => 0x35,
                          'purple'  => 0x14,
                          'red'     => 0x02,
                          'silver'  => 0x16,
                          'white'   => 0x01,
                          'yellow'  => 0x05
                       );

        // Return the default color, 0x7FFF, if undef,
        if ($name_color === '') {
            return(0x7FFF);
        }

        // or the color string converted to an integer,
        if (isset($colors[$name_color])) {
            return($colors[$name_color]);
        }

        // or the default color if string is unrecognised,
        if (preg_match("/\D/",$name_color)) {
            return(0x7FFF);
        }

        // or the default color if arg is outside range,
        if ($name_color > 63) {
            return(0x7FFF);
        }

        // or an integer in the valid range
        return($name_color);
    }

    /**
    * Set cell alignment.
    *
    * @access public
    * @param string $location alignment for the cell ('left', 'right', etc...).
    */
    function setAlign($location)
    {
        $this->setHAlign($location);
        $this->setVAlign($location);
    }

    /**
    * Set cell horizontal alignment.
    *
    * @access public
    * @param string $location alignment for the cell ('left', 'right', etc...).
    */
    function setHAlign($location)
    {
        $location = strtolower((string) $location);

        $map = array(
            'left'          => 1,
            'centre'        => 2,
            'center'        => 2,
            'right'         => 3,
            'fill'          => 4,
            'justify'       => 5,
            'merge'         => 6,
            'equal_space'   => 7,
        );
        if (isset($map[$location])) {
            $this->_text_h_align = $map[$location];
        }
    }

    /**
    * Set cell vertical alignment.
    *
    * @access public
    * @param string $location alignment for the cell ('top', 'vleft', 'vright', etc...).
    */
    function setVAlign($location)
    {
        $location = strtolower((string) $location);

        $map = array(
            'top'           => 0,
            'vcentre'       => 1,
            'vcenter'       => 1,
            'bottom'        => 2,
            'vjustify'      => 3,
            'vequal_space'  => 4,
        );
        if (isset($map[$location])) {
            $this->_text_v_align = $map[$location];
        }
    }

    /**
    * This is an alias for the unintuitive setAlign('merge')
    *
    * @access public
    */
    function setMerge()
    {
        $this->setAlign('merge');
    }

    /**
    * Sets the boldness of the text.
    * Bold has a range 100..1000.
    * 0 (400) is normal. 1 (700) is bold.
    *
    * @access public
    * @param integer $weight Weight for the text, 0 maps to 400 (normal text),
                             1 maps to 700 (bold text). Valid range is: 100-1000.
                             It's Optional, default is 1 (bold).
    */
    function setBold($weight = 1)
    {
        $bold = 400;
        if ($weight == 1) {
            $bold = 700;
        }

        $this->_bold = $bold;
    }


    /************************************
    * FUNCTIONS FOR SETTING CELLS BORDERS
    */

    /**
    * Sets the width for the bottom border of the cell
    *
    * @access public
    * @param integer $style style of the cell border. 1 => thin, 2 => thick.
    */
    function setBottom($style)
    {
        $this->_bottom = $style;
    }

    /**
    * Sets the width for the top border of the cell
    *
    * @access public
    * @param integer $style style of the cell top border. 1 => thin, 2 => thick.
    */
    function setTop($style)
    {
        $this->_top = $style;
    }

    /**
    * Sets the width for the left border of the cell
    *
    * @access public
    * @param integer $style style of the cell left border. 1 => thin, 2 => thick.
    */
    function setLeft($style)
    {
        $this->_left = $style;
    }

    /**
    * Sets the width for the right border of the cell
    *
    * @access public
    * @param integer $style style of the cell right border. 1 => thin, 2 => thick.
    */
    function setRight($style)
    {
        $this->_right = $style;
    }


    /**
    * Set cells borders to the same style
    *
    * @access public
    * @param integer $style style to apply for all cell borders. 1 => thin, 2 => thick.
    */
    function setBorder($style)
    {
        $this->setBottom($style);
        $this->setTop($style);
        $this->setLeft($style);
        $this->setRight($style);
    }


    /*******************************************
    * FUNCTIONS FOR SETTING CELLS BORDERS COLORS
    */

    /**
    * Sets all the cell's borders to the same color
    *
    * @access public
    * @param mixed $color The color we are setting. Either a string (like 'blue'),
    *                     or an integer (range is [8...63]).
    */
    function setBorderColor($color)
    {
        $this->setBottomColor($color);
        $this->setTopColor($color);
        $this->setLeftColor($color);
        $this->setRightColor($color);
    }

    /**
    * Sets the cell's bottom border color
    *
    * @access public
    * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
    */
    function setBottomColor($color)
    {
        $value = $this->_getColor($color);
        $this->_bottom_color = $value;
    }

    /**
    * Sets the cell's top border color
    *
    * @access public
    * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
    */
    function setTopColor($color)
    {
        $value = $this->_getColor($color);
        $this->_top_color = $value;
    }

    /**
    * Sets the cell's left border color
    *
    * @access public
    * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
    */
    function setLeftColor($color)
    {
        $value = $this->_getColor($color);
        $this->_left_color = $value;
    }

    /**
    * Sets the cell's right border color
    *
    * @access public
    * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
    */
    function setRightColor($color)
    {
        $value = $this->_getColor($color);
        $this->_right_color = $value;
    }


    /**
    * Sets the cell's foreground color
    *
    * @access public
    * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
    */
    function setFgColor($color)
    {
        $value = $this->_getColor($color);
        $this->_fg_color = $value;
        if ($this->_pattern == 0) { // force color to be seen
            $this->_pattern = 1;
        }
    }

    /**
    * Sets the cell's background color
    *
    * @access public
    * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
    */
    function setBgColor($color)
    {
        $value = $this->_getColor($color);
        $this->_bg_color = $value;
        if ($this->_pattern == 0) { // force color to be seen
            $this->_pattern = 1;
        }
    }

    /**
    * Sets the cell's color
    *
    * @access public
    * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
    */
    function setColor($color)
    {
        $value = $this->_getColor($color);
        $this->_color = $value;
    }

    /**
    * Sets the fill pattern attribute of a cell
    *
    * @access public
    * @param integer $arg Optional. Defaults to 1. Meaningful values are: 0-18,
    *                     0 meaning no background.
    */
    function setPattern($arg = 1)
    {
        $this->_pattern = $arg;
    }

    /**
    * Sets the underline of the text
    *
    * @access public
    * @param integer $underline The value for underline. Possible values are:
    *                          1 => underline, 2 => double underline.
    */
    function setUnderline($underline)
    {
        $this->_underline = $underline;
    }

    /**
    * Sets the font style as italic
    *
    * @access public
    */
    function setItalic()
    {
        $this->_italic = 1;
    }

    /**
    * Sets the font size
    *
    * @access public
    * @param integer $size The font size (in pixels I think).
    */
    function setSize($size)
    {
        $this->_size = $size;
    }

    /**
    * Sets text wrapping
    *
    * @access public
    */
    function setTextWrap()
    {
        $this->_text_wrap = 1;
    }

    /**
    * Sets the orientation of the text
    *
    * @access public
    * @param integer $angle The rotation angle for the text (clockwise). Possible
                            values are: 0, 90, 270 and -1 for stacking top-to-bottom.
    */
    /*
    function setTextRotation($angle)
    {
        switch ($angle)
        {
            case 0:
                $this->_rotation = 0;
                break;
            case 90:
                $this->_rotation = 3;
                break;
            case 270:
                $this->_rotation = 2;
                break;
            case -1:
                $this->_rotation = 1;
                break;
            default :
                return $this->raiseError("Invalid value for angle.".
                                  " Possible values are: 0, 90, 270 and -1 ".
                                  "for stacking top-to-bottom.");
                $this->_rotation = 0;
                break;
        }
    }
    */

    /**
    * Sets the numeric format.
    * It can be date, time, currency, etc...
    *
    * @access public
    * @param integer $num_format The numeric format.
    */
    function setNumFormat($num_format)
    {
        $this->_num_format = $num_format;
    }

    /**
    * Sets font as strikeout.
    *
    * @access public
    */
    function setStrikeOut()
    {
        $this->_font_strikeout = 1;
    }

    /**
    * Sets outlining for a font.
    *
    * @access public
    */
    function setOutLine()
    {
        $this->_font_outline = 1;
    }

    /**
    * Sets font as shadow.
    *
    * @access public
    */
    function setShadow()
    {
        $this->_font_shadow = 1;
    }

    /**
    * Sets the script type of the text
    *
    * @access public
    * @param integer $script The value for script type. Possible values are:
    *                        1 => superscript, 2 => subscript.
    */
    function setScript($script)
    {
        $this->_font_script = $script;
    }

     /**
     * Locks a cell.
     *
     * @access public
     */
     function setLocked()
     {
         $this->_locked = 1;
     }

    /**
    * Unlocks a cell. Useful for unprotecting particular cells of a protected sheet.
    *
    * @access public
    */
    function setUnLocked()
    {
        $this->_locked = 0;
    }

    /**
    * Sets the font family name.
    *
    * @access public
    * @param string $fontfamily The font family name. Possible values are:
    *                           'Times New Roman', 'Arial', 'Courier'.
    */
    function setFontFamily($font_family)
    {
        $this->_font_name = $font_family;
    }
}
?>
