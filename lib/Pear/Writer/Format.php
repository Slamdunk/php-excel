<?php

namespace Slam\Excel\Pear\Writer;

use Slam\Excel;

/**
 * Class for generating Excel XF records (formats).
 *
 * @author   Xavier Noguer <xnoguer@rezebra.com>
 *
 * @category FileFormats
 */
class Format
{
    /**
     * The index given by the workbook when creating a new format.
     *
     * @var int
     */
    private $_xf_index;

    /**
     * Index to the FONT record.
     *
     * @var int
     */
    public $font_index;

    /**
     * The font name (ASCII).
     *
     * @var string
     */
    private $_font_name;

    /**
     * Height of font (1/20 of a point).
     *
     * @var int
     */
    private $_size;

    /**
     * Bold style.
     *
     * @var int
     */
    private $_bold;

    /**
     * Bit specifiying if the font is italic.
     *
     * @var int
     */
    private $_italic;

    /**
     * Index to the cell's color.
     *
     * @var int
     */
    private $_color;

    /**
     * The text underline property.
     *
     * @var int
     */
    private $_underline;

    /**
     * Bit specifiying if the font has strikeout.
     *
     * @var int
     */
    private $_font_strikeout;

    /**
     * Bit specifiying if the font has outline.
     *
     * @var int
     */
    private $_font_outline;

    /**
     * Bit specifiying if the font has shadow.
     *
     * @var int
     */
    private $_font_shadow;

    /**
     * 2 bytes specifiying the script type for the font.
     *
     * @var int
     */
    private $_font_script;

    /**
     * Byte specifiying the font family.
     *
     * @var int
     */
    private $_font_family;

    /**
     * Byte specifiying the font charset.
     *
     * @var int
     */
    private $_font_charset;

    /**
     * An index (2 bytes) to a FORMAT record (number format).
     *
     * @var int|string
     */
    private $_num_format;

    /**
     * Bit specifying if formulas are hidden.
     *
     * @var int
     */
    private $_hidden;

    /**
     * Bit specifying if the cell is locked.
     *
     * @var int
     */
    private $_locked;

    /**
     * The three bits specifying the text horizontal alignment.
     *
     * @var int
     */
    private $_text_h_align;

    /**
     * Bit specifying if the text is wrapped at the right border.
     *
     * @var int
     */
    private $_text_wrap;

    /**
     * The three bits specifying the text vertical alignment.
     *
     * @var int
     */
    private $_text_v_align;

    /**
     * 1 bit, apparently not used.
     *
     * @var int
     */
    private $_text_justlast;

    /**
     * The two bits specifying the text rotation.
     *
     * @var int
     */
    private $_rotation;

    /**
     * The cell's foreground color.
     *
     * @var int
     */
    private $_fg_color;

    /**
     * The cell's background color.
     *
     * @var int
     */
    private $_bg_color;

    /**
     * The cell's background fill pattern.
     *
     * @var int
     */
    private $_pattern;

    /**
     * Style of the bottom border of the cell.
     *
     * @var int
     */
    private $_bottom;

    /**
     * Color of the bottom border of the cell.
     *
     * @var int
     */
    private $_bottom_color;

    /**
     * Style of the top border of the cell.
     *
     * @var int
     */
    private $_top;

    /**
     * Color of the top border of the cell.
     *
     * @var int
     */
    private $_top_color;

    /**
     * Style of the left border of the cell.
     *
     * @var int
     */
    private $_left;

    /**
     * Color of the left border of the cell.
     *
     * @var int
     */
    private $_left_color;

    /**
     * Style of the right border of the cell.
     *
     * @var int
     */
    private $_right;

    /**
     * Color of the right border of the cell.
     *
     * @var int
     */
    private $_right_color;

    private $_diag;
    private $_diag_color;

    /**
     * Constructor.
     *
     *
     * @param int   $index      the XF index for the format
     * @param array $properties array with properties to be set on initialization
     */
    public function __construct($index = 0, $properties =  [])
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
        foreach ($properties as $property => $value) {
            if (\method_exists($this, 'set' . \ucwords($property))) {
                $method_name = 'set' . \ucwords($property);
                $this->{$method_name}($value);
            }
        }
    }

    /**
     * Generate an Excel BIFF XF record (style or cell).
     *
     * @param string $style the type of the XF record ('style' or 'cell')
     *
     * @return string The XF record
     */
    public function getXf($style)
    {
        // Set the type of the XF record and some of the attributes.
        if ('style' == $style) {
            $style = 0xFFF5;
        } else {
            $style   = $this->_locked;
            $style  |= $this->_hidden << 1;
        }

        // Flags to indicate if attributes have been set.
        $atr_num     = (0 != $this->_num_format) ? 1 : 0;
        $atr_fnt     = (0 != $this->font_index) ? 1 : 0;
        $atr_alc     = ($this->_text_wrap) ? 1 : 0;
        $atr_bdr     = ($this->_bottom   ||
                        $this->_top      ||
                        $this->_left     ||
                        $this->_right) ? 1 : 0;
        $atr_pat     = ((0x40 != $this->_fg_color) ||
                        (0x41 != $this->_bg_color) ||
                        $this->_pattern) ? 1 : 0;
        $atr_prot    = $this->_locked | $this->_hidden;

        // Zero the default border colour if the border has not been set.
        if (0 == $this->_bottom) {
            $this->_bottom_color = 0;
        }
        if (0  == $this->_top) {
            $this->_top_color = 0;
        }
        if (0 == $this->_right) {
            $this->_right_color = 0;
        }
        if (0 == $this->_left) {
            $this->_left_color = 0;
        }
        if (0 == $this->_diag) {
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

        $header      = \pack('vv',       $record, $length);
        $data        = \pack('vvvvvvvv', $ifnt, $ifmt, $style, $align,
                                        $icv, $fill,
                                        $border1, $border2);

        return $header . $data;
    }

    /**
     * Generate an Excel BIFF FONT record.
     *
     * @return string The FONT record
     */
    public function getFont()
    {
        $dyHeight   = $this->_size * 20;    // Height of font (1/20 of a point)
        $icv        = $this->_color;        // Index to color palette
        $bls        = $this->_bold;         // Bold style
        $sss        = $this->_font_script;  // Superscript/subscript
        $uls        = $this->_underline;    // Underline
        $bFamily    = $this->_font_family;  // Font family
        $bCharSet   = $this->_font_charset; // Character set

        $cch        = \strlen($this->_font_name); // Length of font name
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

        $header  = \pack('vv',         $record, $length);
        $data    = \pack('vvvvvCCCCC', $dyHeight, $grbit, $icv, $bls,
                                      $sss, $uls, $bFamily,
                                      $bCharSet, $reserved, $cch);

        return $header . $data . $this->_font_name;
    }

    /**
     * Returns a unique hash key for a font.
     * Used by Excel_Writer_Workbook::_storeAllFonts().
     *
     * The elements that form the key are arranged to increase the probability of
     * generating a unique key. Elements that hold a large range of numbers
     * (eg. _color) are placed between two binary elements such as _italic
     *
     * @return string A key for this font
     */
    public function getFontKey()
    {
        $key  = "$this->_font_name$this->_size";
        $key .= "$this->_font_script$this->_underline";
        $key .= "$this->_font_strikeout$this->_bold$this->_font_outline";
        $key .= "$this->_font_family$this->_font_charset";
        $key .= "$this->_font_shadow$this->_color$this->_italic";
        $key  = \str_replace(' ', '_', $key);

        return $key;
    }

    /**
     * Returns the index used by Excel_Writer_Worksheet::_XF().
     *
     * @return int The index for the XF record
     */
    public function getXfIndex()
    {
        return $this->_xf_index;
    }

    /**
     * Used in conjunction with the set_xxx_color methods to convert a color
     * string into a number. Color range is 0..63 but we will restrict it
     * to 8..63 to comply with Gnumeric. Colors 0..7 are repeated in 8..15.
     *
     *
     * @param mixed $name_color name of the color (i.e.: 'blue', 'red', etc..). Optional.
     *
     * @return int The color index
     */
    private function _getColor($name_color = null)
    {
        $colors = [
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
            'yellow'  => 0x05,
        ];

        // Return the default color, 0x7FFF, if undef,
        if (null === $name_color) {
            return 0x7FFF;
        }

        // or the color string converted to an integer,
        if (isset($colors[$name_color])) {
            return $colors[$name_color];
        }

        // or the default color if string is unrecognised,
        if (\preg_match('/\\D/', $name_color)) {
            return 0x7FFF;
        }

        // or the default color if arg is outside range,
        if ($name_color > 63) {
            return 0x7FFF;
        }

        // or an integer in the valid range
        return $name_color;
    }

    /**
     * Set cell alignment.
     *
     *
     * @param string $location alignment for the cell ('left', 'right', etc...).
     */
    public function setAlign($location)
    {
        $this->setHAlign($location);
        $this->setVAlign($location);
    }

    /**
     * Set cell horizontal alignment.
     *
     *
     * @param mixed $location alignment for the cell ('left', 'right', etc...).
     */
    public function setHAlign($location)
    {
        $location = \strtolower((string) $location);

        $map = [
            'left'          => 1,
            'centre'        => 2,
            'center'        => 2,
            'right'         => 3,
            'fill'          => 4,
            'justify'       => 5,
            'merge'         => 6,
            'equal_space'   => 7,
        ];
        if (isset($map[$location])) {
            $this->_text_h_align = $map[$location];
        }
    }

    /**
     * Set cell vertical alignment.
     *
     *
     * @param mixed $location alignment for the cell ('top', 'vleft', 'vright', etc...).
     */
    public function setVAlign($location)
    {
        $location = \strtolower((string) $location);

        $map = [
            'top'           => 0,
            'vcentre'       => 1,
            'vcenter'       => 1,
            'bottom'        => 2,
            'vjustify'      => 3,
            'vequal_space'  => 4,
        ];
        if (isset($map[$location])) {
            $this->_text_v_align = $map[$location];
        }
    }

    /**
     * This is an alias for the unintuitive setAlign('merge').
     */
    public function setMerge()
    {
        $this->setAlign('merge');
    }

    /**
     * Sets the boldness of the text.
     * Bold has a range 100..1000.
     * 0 (400) is normal. 1 (700) is bold.
     *
     *
     * @param int $weight Weight for the text, 0 maps to 400 (normal text),
     */
    public function setBold($weight = 1)
    {
        $bold = 400;
        if (1 == $weight) {
            $bold = 700;
        }

        $this->_bold = $bold;
    }

    // FUNCTIONS FOR SETTING CELLS BORDERS

    /**
     * Sets the width for the bottom border of the cell.
     *
     *
     * @param int $style style of the cell border. 1 => thin, 2 => thick.
     */
    public function setBottom($style)
    {
        $this->_bottom = $style;
    }

    /**
     * Sets the width for the top border of the cell.
     *
     *
     * @param int $style style of the cell top border. 1 => thin, 2 => thick.
     */
    public function setTop($style)
    {
        $this->_top = $style;
    }

    /**
     * Sets the width for the left border of the cell.
     *
     *
     * @param int $style style of the cell left border. 1 => thin, 2 => thick.
     */
    public function setLeft($style)
    {
        $this->_left = $style;
    }

    /**
     * Sets the width for the right border of the cell.
     *
     *
     * @param int $style style of the cell right border. 1 => thin, 2 => thick.
     */
    public function setRight($style)
    {
        $this->_right = $style;
    }

    /**
     * Set cells borders to the same style.
     *
     *
     * @param int $style style to apply for all cell borders. 1 => thin, 2 => thick.
     */
    public function setBorder($style)
    {
        $this->setBottom($style);
        $this->setTop($style);
        $this->setLeft($style);
        $this->setRight($style);
    }

    // FUNCTIONS FOR SETTING CELLS BORDERS COLORS

    /**
     * Sets all the cell's borders to the same color.
     *
     *
     * @param mixed $color The color we are setting. Either a string (like 'blue'),
     *                     or an integer (range is [8...63]).
     */
    public function setBorderColor($color)
    {
        $this->setBottomColor($color);
        $this->setTopColor($color);
        $this->setLeftColor($color);
        $this->setRightColor($color);
    }

    /**
     * Sets the cell's bottom border color.
     *
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setBottomColor($color)
    {
        $value = $this->_getColor($color);
        $this->_bottom_color = $value;
    }

    /**
     * Sets the cell's top border color.
     *
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setTopColor($color)
    {
        $value = $this->_getColor($color);
        $this->_top_color = $value;
    }

    /**
     * Sets the cell's left border color.
     *
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setLeftColor($color)
    {
        $value = $this->_getColor($color);
        $this->_left_color = $value;
    }

    /**
     * Sets the cell's right border color.
     *
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setRightColor($color)
    {
        $value = $this->_getColor($color);
        $this->_right_color = $value;
    }

    /**
     * Sets the cell's foreground color.
     *
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setFgColor($color)
    {
        $value = $this->_getColor($color);
        $this->_fg_color = $value;
        if (0 == $this->_pattern) { // force color to be seen
            $this->_pattern = 1;
        }
    }

    /**
     * Sets the cell's background color.
     *
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setBgColor($color)
    {
        $value = $this->_getColor($color);
        $this->_bg_color = $value;
        if (0 == $this->_pattern) { // force color to be seen
            $this->_pattern = 1;
        }
    }

    /**
     * Sets the cell's color.
     *
     *
     * @param mixed $color either a string (like 'blue'), or an integer (range is [8...63]).
     */
    public function setColor($color)
    {
        $value = $this->_getColor($color);
        $this->_color = $value;
    }

    /**
     * Sets the fill pattern attribute of a cell.
     *
     *
     * @param int $arg Optional. Defaults to 1. Meaningful values are: 0-18,
     *                 0 meaning no background.
     */
    public function setPattern($arg = 1)
    {
        $this->_pattern = $arg;
    }

    /**
     * Sets the underline of the text.
     *
     *
     * @param int $underline The value for underline. Possible values are:
     *                       1 => underline, 2 => double underline.
     */
    public function setUnderline($underline)
    {
        $this->_underline = $underline;
    }

    /**
     * Sets the font style as italic.
     */
    public function setItalic()
    {
        $this->_italic = 1;
    }

    /**
     * Sets the font size.
     *
     *
     * @param int $size the font size (in pixels I think)
     */
    public function setSize($size)
    {
        $this->_size = $size;
    }

    /**
     * Sets text wrapping.
     */
    public function setTextWrap()
    {
        $this->_text_wrap = 1;
    }

    /**
     * Sets the orientation of the text.
     *
     *
     * @param int $angle The rotation angle for the text (clockwise). Possible
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
                throw new Excel\Exception\RuntimeException("Invalid value for angle. Possible values are: 0, 90, 270 and -1 for stacking top-to-bottom.");
                break;
        }
    }
    */

    /**
     * Sets the numeric format.
     * It can be date, time, currency, etc...
     *
     *
     * @param int|string $num_format the numeric format
     */
    public function setNumFormat($num_format)
    {
        $this->_num_format = $num_format;
    }

    public function getNumFormat()
    {
        return $this->_num_format;
    }

    /**
     * Sets font as strikeout.
     */
    public function setStrikeOut()
    {
        $this->_font_strikeout = 1;
    }

    /**
     * Sets outlining for a font.
     */
    public function setOutLine()
    {
        $this->_font_outline = 1;
    }

    /**
     * Sets font as shadow.
     */
    public function setShadow()
    {
        $this->_font_shadow = 1;
    }

    /**
     * Sets the script type of the text.
     *
     *
     * @param int $script The value for script type. Possible values are:
     *                    1 => superscript, 2 => subscript.
     */
    public function setScript($script)
    {
        $this->_font_script = $script;
    }

    /**
     * Locks a cell.
     */
    public function setLocked()
    {
        $this->_locked = 1;
    }

    /**
     * Unlocks a cell. Useful for unprotecting particular cells of a protected sheet.
     */
    public function setUnLocked()
    {
        $this->_locked = 0;
    }

    /**
     * Sets the font family name.
     *
     *
     * @param string $font_family The font family name. Possible values are:
     *                            'Times New Roman', 'Arial', 'Courier'.
     */
    public function setFontFamily($font_family)
    {
        $this->_font_name = $font_family;
    }
}
