<?php

namespace Slam\Excel\Pear\Writer;

use Slam\Excel;

/**
 * Class for generating Excel Spreadsheets.
 *
 * @author   Xavier Noguer <xnoguer@rezebra.com>
 *
 * @category FileFormats
 */
class Workbook extends BIFFwriter
{
    /**
     * Filename for the Workbook.
     *
     * @var string
     */
    protected $_filename;

    /**
     * Formula parser.
     *
     * @var Parser
     */
    protected $_parser;

    /**
     * Flag for 1904 date system (0 => base date is 1900, 1 => base date is 1904).
     *
     * @var int
     */
    protected $_1904;

    /**
     * The active worksheet of the workbook (0 indexed).
     *
     * @var int
     */
    protected $_activesheet;

    /**
     * 1st displayed worksheet in the workbook (0 indexed).
     *
     * @var int
     */
    protected $_firstsheet;

    /**
     * Number of workbook tabs selected.
     *
     * @var int
     */
    protected $_selected;

    /**
     * Index for creating adding new formats to the workbook.
     *
     * @var int
     */
    protected $_xf_index;

    /**
     * Flag for preventing close from being called twice.
     *
     * @var int
     *
     * @see close()
     */
    protected $_fileclosed;

    /**
     * The BIFF file size for the workbook.
     *
     * @var int
     *
     * @see _calcSheetOffsets()
     */
    protected $_biffsize;

    /**
     * The default sheetname for all sheets created.
     *
     * @var string
     */
    protected $_sheetname;

    /**
     * The default XF format.
     *
     * @var Format
     */
    protected $_tmp_format;

    /**
     * Array containing references to all of this workbook's worksheets.
     *
     * @var array
     */
    protected $_worksheets;

    /**
     * Array of sheetnames for creating the EXTERNSHEET records.
     *
     * @var array
     */
    protected $_sheetnames;

    /**
     * Array containing references to all of this workbook's formats.
     *
     * @var array
     */
    protected $_formats;

    /**
     * Array containing the colour palette.
     *
     * @var array
     */
    protected $_palette;

    /**
     * The default format for URLs.
     *
     * @var Format
     */
    protected $_url_format;

    /**
     * The codepage indicates the text encoding used for strings.
     *
     * @var int
     */
    protected $_codepage;

    /**
     * The country code used for localization.
     *
     * @var int
     */
    protected $_country_code;

    protected $_str_total;
    protected $_str_unique;
    protected $_str_table;
    protected $_block_sizes;

    /**
     * Class constructor.
     *
     * @param string $filename for storing the workbook. "-" for writing to stdout.
     */
    public function __construct($filename)
    {
        parent::__construct();

        $this->_filename            = $filename;
        $this->_parser              = new Parser($this->_byte_order);
        $this->_1904                = 0;
        $this->_activesheet         = 0;
        $this->_firstsheet          = 0;
        $this->_selected            = 0;
        $this->_xf_index            = 16; // 15 style XF's and 1 cell XF.
        $this->_fileclosed          = 0;
        $this->_biffsize            = 0;
        $this->_sheetname           = 'Sheet';
        $this->_tmp_format          = new Format();
        $this->_worksheets          = [];
        $this->_sheetnames          = [];
        $this->_formats             = [];
        $this->_palette             = [];
        $this->_codepage            = 0x04E4; // FIXME: should change for BIFF8
        $this->_country_code        = -1;

        // Add the default format for hyperlinks
        $this->_url_format          = $this->addFormat(['color' => 'blue', 'underline' => 1]);
        $this->_str_total           = 0;
        $this->_str_unique          = 0;
        $this->_str_table           = [];
        $this->_setPaletteXl97();
    }

    /**
     * Calls finalization methods.
     * This method should always be the last one to be called on every workbook.
     *
     *
     * @return mixed true on success. Excel_PEAR_Error on failure
     */
    public function close()
    {
        if ($this->_fileclosed) { // Prevent close() from being called twice.
            return true;
        }
        $this->_storeWorkbook();

        $this->_fileclosed = 1;

        foreach ($this->_worksheets as $sheet) {
            $sheet->fclose();
        }

        return true;
    }

    /**
     * An accessor for the _worksheets[] array
     * Returns an array of the worksheet objects in a workbook
     * It actually calls to worksheets().
     *
     *
     * @see worksheets()
     *
     * @return array
     */
    public function sheets()
    {
        return $this->worksheets();
    }

    /**
     * An accessor for the _worksheets[] array.
     * Returns an array of the worksheet objects in a workbook.
     *
     *
     * @return array
     */
    public function worksheets()
    {
        return $this->_worksheets;
    }

    /**
     * Set the country identifier for the workbook.
     *
     *
     * @param int $code is the international calling country code for the
     *                  chosen country
     */
    public function setCountry($code)
    {
        $this->_country_code = $code;
    }

    /**
     * Add a new worksheet to the Excel workbook.
     * If no name is given the name of the worksheet will be Sheeti$i, with
     * $i in [1..].
     *
     *
     * @param string $name the optional name of the worksheet
     *
     * @return mixed reference to a worksheet object on success, Excel_PEAR_Error
     *               on failure
     */
    public function addWorksheet($name = '')
    {
        $index     = \count($this->_worksheets);
        $sheetname = $this->_sheetname;

        if ('' == $name) {
            $name = $sheetname . ($index + 1);
        }

        // Check that sheetname is <= 31 chars (Excel limit before BIFF8).
        if (\strlen($name) > 31) {
            throw new Excel\Exception\RuntimeException("Sheetname ${name} must be <= 31 chars");
        }

        // Check that the worksheet name doesn't already exist: a fatal Excel error.
        $total_worksheets = \count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            if ($this->_worksheets[$i]->getName() == $name) {
                throw new Excel\Exception\RuntimeException("Worksheet '${name}' already exists");
            }
        }

        $worksheet = new Worksheet(
            $name, $index,
            $this->_activesheet, $this->_firstsheet,
            $this->_str_total, $this->_str_unique,
            $this->_str_table, $this->_url_format,
            $this->_parser
        );

        $this->_worksheets[$index] = $worksheet;     // Store ref for iterator
        $this->_sheetnames[$index] = $name;          // Store EXTERNSHEET names
        $this->_parser->setExtSheet($name, $index);  // Register worksheet name with parser

        return $worksheet;
    }

    /**
     * Add a new format to the Excel workbook.
     * Also, pass any properties to the Format constructor.
     *
     *
     * @param array $properties array with properties for initializing the format
     *
     * @return Format reference to an Excel Format
     */
    public function addFormat($properties = [])
    {
        $format = new Format($this->_xf_index, $properties);
        $this->_xf_index += 1;
        $this->_formats[] = $format;

        return $format;
    }

    /**
     * Change the RGB components of the elements in the colour palette.
     *
     *
     * @param int $index colour index
     * @param int $red   red RGB value [0-255]
     * @param int $green green RGB value [0-255]
     * @param int $blue  blue RGB value [0-255]
     *
     * @return int The palette index for the custom color
     */
    public function setCustomColor($index, $red, $green, $blue)
    {
        // Match a HTML #xxyyzz style parameter
        /*if (defined $_[1] and $_[1] =~ /^#(\w\w)(\w\w)(\w\w)/ ) {
            @_ = ($_[0], hex $1, hex $2, hex $3);
        }*/

        // Check that the colour index is the right range
        if ($index < 8 or $index > 64) {
            // TODO: assign real error codes
            throw new Excel\Exception\RuntimeException("Color index ${index} outside range: 8 <= index <= 64");
        }

        // Check that the colour components are in the right range
        if (($red   < 0 or $red   > 255) ||
            ($green < 0 or $green > 255) ||
            ($blue  < 0 or $blue  > 255)) {
            throw new Excel\Exception\RuntimeException('Color component outside range: 0 <= color <= 255');
        }

        $index -= 8; // Adjust colour index (wingless dragonfly)

        // Set the RGB value
        $this->_palette[$index] = [$red, $green, $blue, 0];

        return $index + 8;
    }

    /**
     * Sets the colour palette to the Excel 97+ default.
     */
    protected function _setPaletteXl97()
    {
        $this->_palette = [
            [0x00, 0x00, 0x00, 0x00],   // 8
            [0xff, 0xff, 0xff, 0x00],   // 9
            [0xff, 0x00, 0x00, 0x00],   // 10
            [0x00, 0xff, 0x00, 0x00],   // 11
            [0x00, 0x00, 0xff, 0x00],   // 12
            [0xff, 0xff, 0x00, 0x00],   // 13
            [0xff, 0x00, 0xff, 0x00],   // 14
            [0x00, 0xff, 0xff, 0x00],   // 15
            [0x80, 0x00, 0x00, 0x00],   // 16
            [0x00, 0x80, 0x00, 0x00],   // 17
            [0x00, 0x00, 0x80, 0x00],   // 18
            [0x80, 0x80, 0x00, 0x00],   // 19
            [0x80, 0x00, 0x80, 0x00],   // 20
            [0x00, 0x80, 0x80, 0x00],   // 21
            [0xc0, 0xc0, 0xc0, 0x00],   // 22
            [0x80, 0x80, 0x80, 0x00],   // 23
            [0x99, 0x99, 0xff, 0x00],   // 24
            [0x99, 0x33, 0x66, 0x00],   // 25
            [0xff, 0xff, 0xcc, 0x00],   // 26
            [0xcc, 0xff, 0xff, 0x00],   // 27
            [0x66, 0x00, 0x66, 0x00],   // 28
            [0xff, 0x80, 0x80, 0x00],   // 29
            [0x00, 0x66, 0xcc, 0x00],   // 30
            [0xcc, 0xcc, 0xff, 0x00],   // 31
            [0x00, 0x00, 0x80, 0x00],   // 32
            [0xff, 0x00, 0xff, 0x00],   // 33
            [0xff, 0xff, 0x00, 0x00],   // 34
            [0x00, 0xff, 0xff, 0x00],   // 35
            [0x80, 0x00, 0x80, 0x00],   // 36
            [0x80, 0x00, 0x00, 0x00],   // 37
            [0x00, 0x80, 0x80, 0x00],   // 38
            [0x00, 0x00, 0xff, 0x00],   // 39
            [0x00, 0xcc, 0xff, 0x00],   // 40
            [0xcc, 0xff, 0xff, 0x00],   // 41
            [0xcc, 0xff, 0xcc, 0x00],   // 42
            [0xff, 0xff, 0x99, 0x00],   // 43
            [0x99, 0xcc, 0xff, 0x00],   // 44
            [0xff, 0x99, 0xcc, 0x00],   // 45
            [0xcc, 0x99, 0xff, 0x00],   // 46
            [0xff, 0xcc, 0x99, 0x00],   // 47
            [0x33, 0x66, 0xff, 0x00],   // 48
            [0x33, 0xcc, 0xcc, 0x00],   // 49
            [0x99, 0xcc, 0x00, 0x00],   // 50
            [0xff, 0xcc, 0x00, 0x00],   // 51
            [0xff, 0x99, 0x00, 0x00],   // 52
            [0xff, 0x66, 0x00, 0x00],   // 53
            [0x66, 0x66, 0x99, 0x00],   // 54
            [0x96, 0x96, 0x96, 0x00],   // 55
            [0x00, 0x33, 0x66, 0x00],   // 56
            [0x33, 0x99, 0x66, 0x00],   // 57
            [0x00, 0x33, 0x00, 0x00],   // 58
            [0x33, 0x33, 0x00, 0x00],   // 59
            [0x99, 0x33, 0x00, 0x00],   // 60
            [0x99, 0x33, 0x66, 0x00],   // 61
            [0x33, 0x33, 0x99, 0x00],   // 62
            [0x33, 0x33, 0x33, 0x00],   // 63
         ];
    }

    /**
     * Assemble worksheets into a workbook and send the BIFF data to an Excel_OLE
     * storage.
     *
     *
     * @return mixed true on success. Excel_PEAR_Error on failure
     */
    protected function _storeWorkbook()
    {
        if (0 == \count($this->_worksheets)) {
            return true;
        }

        // Ensure that at least one worksheet has been selected.
        if (0 == $this->_activesheet) {
            $this->_worksheets[0]->selected = 1;
        }

        // Calculate the number of selected worksheet tabs and call the finalization
        // methods for each worksheet
        $total_worksheets = \count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            if ($this->_worksheets[$i]->selected) {
                ++$this->_selected;
            }
            $this->_worksheets[$i]->close($this->_sheetnames);
        }

        // Add Workbook globals
        $this->_storeBof(0x0005);
        $this->_storeCodepage();
        $this->_storeExterns();    // For print area and repeat rows
        $this->_storeNames();      // For print area and repeat rows
        $this->_storeWindow1();
        $this->_storeDatemode();
        $this->_storeAllFonts();
        $this->_storeAllNumFormats();
        $this->_storeAllXfs();
        $this->_storeAllStyles();
        $this->_storePalette();
        $this->_calcSheetOffsets();

        // Add BOUNDSHEET records
        for ($i = 0; $i < $total_worksheets; ++$i) {
            $this->_storeBoundsheet($this->_worksheets[$i]->name, $this->_worksheets[$i]->offset);
        }

        if ($this->_country_code != -1) {
            $this->_storeCountry();
        }

        // End Workbook globals
        $this->_storeEof();

        // Store the workbook in an Excel_OLE container
        $this->_storeExcel_OLEFile();

        return true;
    }

    /**
     * Store the workbook in an Excel_OLE container.
     *
     *
     * @return mixed true on success. Excel_PEAR_Error on failure
     */
    protected function _storeExcel_OLEFile()
    {
        $Excel_OLE = new Excel\Pear\OLE\PPS\File(Excel\Pear\OLE::Asc2Ucs('Book'));
        $Excel_OLE->append($this->_data);

        $total_worksheets = \count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            while ($tmp = $this->_worksheets[$i]->getData()) {
                $Excel_OLE->append($tmp);
            }
        }

        $root = new Excel\Pear\OLE\PPS\Root(\time(), \time(), [$Excel_OLE]);
        $root->save($this->_filename);

        return true;
    }

    /**
     * Calculate offsets for Worksheet BOF records.
     */
    protected function _calcSheetOffsets()
    {
        $boundsheet_length = 11;
        $EOF               = 4;
        $offset            = $this->_datasize;

        $total_worksheets = \count($this->_worksheets);
        // add the length of the BOUNDSHEET records
        for ($i = 0; $i < $total_worksheets; ++$i) {
            $offset += $boundsheet_length + \strlen($this->_worksheets[$i]->name);
        }
        $offset += $EOF;

        for ($i = 0; $i < $total_worksheets; ++$i) {
            $this->_worksheets[$i]->offset = $offset;
            $offset += $this->_worksheets[$i]->_datasize;
        }
        $this->_biffsize = $offset;
    }

    /**
     * Store the Excel FONT records.
     */
    protected function _storeAllFonts()
    {
        // tmp_format is added by the constructor. We use this to write the default XF's
        $format = $this->_tmp_format;
        $font   = $format->getFont();

        // Note: Fonts are 0-indexed. According to the SDK there is no index 4,
        // so the following fonts are 0, 1, 2, 3, 5
        //
        for ($i = 1; $i <= 5; ++$i) {
            $this->_append($font);
        }

        // Iterate through the XF objects and write a FONT record if it isn't the
        // same as the default FONT and if it hasn't already been used.
        //
        $fonts = [];
        $index = 6;                  // The first user defined FONT

        $key = $format->getFontKey(); // The default font from _tmp_format
        $fonts[$key] = 0;             // Index of the default font

        $total_formats = \count($this->_formats);
        for ($i = 0; $i < $total_formats; ++$i) {
            $key = $this->_formats[$i]->getFontKey();
            if (isset($fonts[$key])) {
                // FONT has already been used
                $this->_formats[$i]->font_index = $fonts[$key];
            } else {
                // Add a new FONT record
                $fonts[$key]        = $index;
                $this->_formats[$i]->font_index = $index;
                ++$index;
                $font = $this->_formats[$i]->getFont();
                $this->_append($font);
            }
        }
    }

    /**
     * Store user defined numerical formats i.e. FORMAT records.
     */
    protected function _storeAllNumFormats()
    {
        // Leaning num_format syndrome
        $hash_num_formats = [];
        $num_formats      = [];
        $index = 164;

        // Iterate through the XF objects and write a FORMAT record if it isn't a
        // built-in format type and if the FORMAT string hasn't already been used.
        $total_formats = \count($this->_formats);
        for ($i = 0; $i < $total_formats; ++$i) {
            $num_format = $this->_formats[$i]->getNumFormat();

            // Check if $num_format is an index to a built-in format.
            // Also check for a string of zeros, which is a valid format string
            // but would evaluate to zero.
            //
            if (! \preg_match('/^0+\\d/', $num_format)) {
                if (\preg_match('/^\\d+$/', $num_format)) { // built-in format
                    continue;
                }
            }

            if (isset($hash_num_formats[$num_format])) {
                // FORMAT has already been used
                $this->_formats[$i]->setNumFormat($hash_num_formats[$num_format]);
            } else {
                // Add a new FORMAT
                $hash_num_formats[$num_format]  = $index;
                $this->_formats[$i]->setNumFormat($index);
                \array_push($num_formats, $num_format);
                ++$index;
            }
        }

        // Write the new FORMAT records starting from 0xA4
        $index = 164;
        foreach ($num_formats as $num_format) {
            $this->_storeNumFormat($num_format, $index);
            ++$index;
        }
    }

    /**
     * Write all XF records.
     */
    protected function _storeAllXfs()
    {
        // _tmp_format is added by the constructor. We use this to write the default XF's
        // The default font index is 0
        //
        $format = $this->_tmp_format;
        for ($i = 0; $i <= 14; ++$i) {
            $xf = $format->getXf('style'); // Style XF
            $this->_append($xf);
        }

        $xf = $format->getXf('cell');      // Cell XF
        $this->_append($xf);

        // User defined XFs
        $total_formats = \count($this->_formats);
        for ($i = 0; $i < $total_formats; ++$i) {
            $xf = $this->_formats[$i]->getXf('cell');
            $this->_append($xf);
        }
    }

    /**
     * Write all STYLE records.
     */
    protected function _storeAllStyles()
    {
        $this->_storeStyle();
    }

    /**
     * Write the EXTERNCOUNT and EXTERNSHEET records. These are used as indexes for
     * the NAME records.
     */
    protected function _storeExterns()
    {
        // Create EXTERNCOUNT with number of worksheets
        $this->_storeExterncount(\count($this->_worksheets));

        // Create EXTERNSHEET for each worksheet
        foreach ($this->_sheetnames as $sheetname) {
            $this->_storeExternsheet($sheetname);
        }
    }

    /**
     * Write the NAME record to define the print area and the repeat rows and cols.
     */
    protected function _storeNames()
    {
        // Create the print area NAME records
        $total_worksheets = \count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            // Write a Name record if the print area has been defined
            if (isset($this->_worksheets[$i]->print_rowmin)) {
                $this->_storeNameShort(
                    $this->_worksheets[$i]->index,
                    0x06, // NAME type
                    $this->_worksheets[$i]->print_rowmin,
                    $this->_worksheets[$i]->print_rowmax,
                    $this->_worksheets[$i]->print_colmin,
                    $this->_worksheets[$i]->print_colmax
                    );
            }
        }

        // Create the print title NAME records
        $total_worksheets = \count($this->_worksheets);
        for ($i = 0; $i < $total_worksheets; ++$i) {
            $rowmin = $this->_worksheets[$i]->title_rowmin;
            $rowmax = $this->_worksheets[$i]->title_rowmax;
            $colmin = $this->_worksheets[$i]->title_colmin;
            $colmax = $this->_worksheets[$i]->title_colmax;

            // Determine if row + col, row, col or nothing has been defined
            // and write the appropriate record
            //
            if (isset($rowmin) && isset($colmin)) {
                // Row and column titles have been defined.
                // Row title has been defined.
                $this->_storeNameLong(
                    $this->_worksheets[$i]->index,
                    0x07, // NAME type
                    $rowmin,
                    $rowmax,
                    $colmin,
                    $colmax
                );
            } elseif (isset($rowmin)) {
                // Row title has been defined.
                $this->_storeNameShort(
                    $this->_worksheets[$i]->index,
                    0x07, // NAME type
                    $rowmin,
                    $rowmax,
                    0x00,
                    0xff
                );
            } elseif (isset($colmin)) {
                // Column title has been defined.
                $this->_storeNameShort(
                    $this->_worksheets[$i]->index,
                    0x07, // NAME type
                    0x0000,
                    0x3fff,
                    $colmin,
                    $colmax
                );
            }
            // Print title hasn't been defined.
        }
    }

    // BIFF RECORDS

    /**
     * Stores the CODEPAGE biff record.
     */
    protected function _storeCodepage()
    {
        $record          = 0x0042;             // Record identifier
        $length          = 0x0002;             // Number of bytes to follow
        $cv              = $this->_codepage;   // The code page

        $header          = \pack('vv', $record, $length);
        $data            = \pack('v',  $cv);

        $this->_append($header . $data);
    }

    /**
     * Write Excel BIFF WINDOW1 record.
     */
    protected function _storeWindow1()
    {
        $record    = 0x003D;                 // Record identifier
        $length    = 0x0012;                 // Number of bytes to follow

        $xWn       = 0x0000;                 // Horizontal position of window
        $yWn       = 0x0000;                 // Vertical position of window
        $dxWn      = 0x25BC;                 // Width of window
        $dyWn      = 0x1572;                 // Height of window

        $grbit     = 0x0038;                 // Option flags
        $ctabsel   = $this->_selected;       // Number of workbook tabs selected
        $wTabRatio = 0x0258;                 // Tab to scrollbar ratio

        $itabFirst = $this->_firstsheet;     // 1st displayed worksheet
        $itabCur   = $this->_activesheet;    // Active worksheet

        $header    = \pack('vv',        $record, $length);
        $data      = \pack('vvvvvvvvv', $xWn, $yWn, $dxWn, $dyWn,
                                       $grbit,
                                       $itabCur, $itabFirst,
                                       $ctabsel, $wTabRatio
        );
        $this->_append($header . $data);
    }

    /**
     * Writes Excel BIFF BOUNDSHEET record.
     * FIXME: inconsistent with BIFF documentation.
     *
     * @param string $sheetname Worksheet name
     * @param int    $offset    Location of worksheet BOF
     */
    protected function _storeBoundsheet($sheetname, $offset)
    {
        $record    = 0x0085;                    // Record identifier
        $length = 0x07 + \strlen($sheetname); // Number of bytes to follow

        $grbit     = 0x0000;                    // Visibility and sheet type
        $cch       = \strlen($sheetname);        // Length of sheet name

        $header    = \pack('vv',  $record, $length);
        $data      = \pack('VvC', $offset, $grbit, $cch);

        $this->_append($header . $data . $sheetname);
    }

    /**
     * Write Internal SUPBOOK record.
     */
    protected function _storeSupbookInternal()
    {
        $record    = 0x01AE;   // Record identifier
        $length    = 0x0004;   // Bytes to follow

        $header    = \pack('vv', $record, $length);
        $data      = \pack('vv', \count($this->_worksheets), 0x0104);
        $this->_append($header . $data);
    }

    /**
     * Writes the Excel BIFF EXTERNSHEET record. These references are used by
     * formulas.
     */
    /*
    protected function _storeExternsheetBiff8()
    {
        $total_references = count($this->_parser->_references);
        $record   = 0x0017;                     // Record identifier
        $length   = 2 + 6 * $total_references;  // Number of bytes to follow

        $header           = pack('vv',  $record, $length);
        $data             = pack('v', $total_references);
        for ($i = 0; $i < $total_references; ++$i) {
            $data .= $this->_parser->_references[$i];
        }
        $this->_append($header . $data);
    }
     */

    /**
     * Write Excel BIFF STYLE records.
     */
    protected function _storeStyle()
    {
        $record    = 0x0293;   // Record identifier
        $length    = 0x0004;   // Bytes to follow

        $ixfe      = 0x8000;   // Index to style XF
        $BuiltIn   = 0x00;     // Built-in style
        $iLevel    = 0xff;     // Outline style level

        $header    = \pack('vv',  $record, $length);
        $data      = \pack('vCC', $ixfe, $BuiltIn, $iLevel);
        $this->_append($header . $data);
    }

    /**
     * Writes Excel FORMAT record for non "built-in" numerical formats.
     *
     * @param string $format Custom format string
     * @param int    $ifmt   Format index code
     */
    protected function _storeNumFormat($format, $ifmt)
    {
        $record    = 0x041E;                      // Record identifier

        $length = \strlen($format);

        $header    = \pack('vv', $record, 3 + $length);
        $data      = \pack('vC', $ifmt, $length);

        $this->_append($header . $data . $format);
    }

    /**
     * Write DATEMODE record to indicate the date system in use (1904 or 1900).
     */
    protected function _storeDatemode()
    {
        $record    = 0x0022;         // Record identifier
        $length    = 0x0002;         // Bytes to follow

        $f1904     = $this->_1904;   // Flag for 1904 date system

        $header    = \pack('vv', $record, $length);
        $data      = \pack('v', $f1904);
        $this->_append($header . $data);
    }

    /**
     * Write BIFF record EXTERNCOUNT to indicate the number of external sheet
     * references in the workbook.
     *
     * Excel only stores references to external sheets that are used in NAME.
     * The workbook NAME record is required to define the print area and the repeat
     * rows and columns.
     *
     * A similar method is used in Worksheet.php for a slightly different purpose.
     *
     * @param int $cxals Number of external references
     */
    protected function _storeExterncount($cxals)
    {
        $record   = 0x0016;          // Record identifier
        $length   = 0x0002;          // Number of bytes to follow

        $header   = \pack('vv', $record, $length);
        $data     = \pack('v',  $cxals);
        $this->_append($header . $data);
    }

    /**
     * Writes the Excel BIFF EXTERNSHEET record. These references are used by
     * formulas. NAME record is required to define the print area and the repeat
     * rows and columns.
     *
     * A similar method is used in Worksheet.php for a slightly different purpose.
     *
     * @param string $sheetname Worksheet name
     */
    protected function _storeExternsheet($sheetname)
    {
        $record      = 0x0017;                     // Record identifier
        $length      = 0x02 + \strlen($sheetname);  // Number of bytes to follow

        $cch         = \strlen($sheetname);         // Length of sheet name
        $rgch        = 0x03;                       // Filename encoding

        $header      = \pack('vv',  $record, $length);
        $data        = \pack('CC', $cch, $rgch);
        $this->_append($header . $data . $sheetname);
    }

    /**
     * Store the NAME record in the short format that is used for storing the print
     * area, repeat rows only and repeat columns only.
     *
     * @param int $index  Sheet index
     * @param int $type   Built-in name type
     * @param int $rowmin Start row
     * @param int $rowmax End row
     * @param int $colmin Start colum
     * @param int $colmax End column
     */
    protected function _storeNameShort($index, $type, $rowmin, $rowmax, $colmin, $colmax)
    {
        $record          = 0x0018;       // Record identifier
        $length          = 0x0024;       // Number of bytes to follow

        $grbit           = 0x0020;       // Option flags
        $chKey           = 0x00;         // Keyboard shortcut
        $cch             = 0x01;         // Length of text name
        $cce             = 0x0015;       // Length of text definition
        $ixals           = $index + 1;   // Sheet index
        $itab            = $ixals;       // Equal to ixals
        $cchCustMenu     = 0x00;         // Length of cust menu text
        $cchDescription  = 0x00;         // Length of description text
        $cchHelptopic    = 0x00;         // Length of help topic text
        $cchStatustext   = 0x00;         // Length of status bar text
        $rgch            = $type;        // Built-in name type

        $unknown03       = 0x3b;
        $unknown04       = 0xffff - $index;
        $unknown05       = 0x0000;
        $unknown06       = 0x0000;
        $unknown07       = 0x1087;
        $unknown08       = 0x8005;

        $header             = \pack('vv', $record, $length);
        $data               = \pack('v', $grbit);
        $data              .= \pack('C', $chKey);
        $data              .= \pack('C', $cch);
        $data              .= \pack('v', $cce);
        $data              .= \pack('v', $ixals);
        $data              .= \pack('v', $itab);
        $data              .= \pack('C', $cchCustMenu);
        $data              .= \pack('C', $cchDescription);
        $data              .= \pack('C', $cchHelptopic);
        $data              .= \pack('C', $cchStatustext);
        $data              .= \pack('C', $rgch);
        $data              .= \pack('C', $unknown03);
        $data              .= \pack('v', $unknown04);
        $data              .= \pack('v', $unknown05);
        $data              .= \pack('v', $unknown06);
        $data              .= \pack('v', $unknown07);
        $data              .= \pack('v', $unknown08);
        $data              .= \pack('v', $index);
        $data              .= \pack('v', $index);
        $data              .= \pack('v', $rowmin);
        $data              .= \pack('v', $rowmax);
        $data              .= \pack('C', $colmin);
        $data              .= \pack('C', $colmax);
        $this->_append($header . $data);
    }

    /**
     * Store the NAME record in the long format that is used for storing the repeat
     * rows and columns when both are specified. This shares a lot of code with
     * _storeNameShort() but we use a separate method to keep the code clean.
     * Code abstraction for reuse can be carried too far, and I should know. ;-).
     *
     * @param int $index  Sheet index
     * @param int $type   Built-in name type
     * @param int $rowmin Start row
     * @param int $rowmax End row
     * @param int $colmin Start colum
     * @param int $colmax End column
     */
    protected function _storeNameLong($index, $type, $rowmin, $rowmax, $colmin, $colmax)
    {
        $record          = 0x0018;       // Record identifier
        $length          = 0x003d;       // Number of bytes to follow
        $grbit           = 0x0020;       // Option flags
        $chKey           = 0x00;         // Keyboard shortcut
        $cch             = 0x01;         // Length of text name
        $cce             = 0x002e;       // Length of text definition
        $ixals           = $index + 1;   // Sheet index
        $itab            = $ixals;       // Equal to ixals
        $cchCustMenu     = 0x00;         // Length of cust menu text
        $cchDescription  = 0x00;         // Length of description text
        $cchHelptopic    = 0x00;         // Length of help topic text
        $cchStatustext   = 0x00;         // Length of status bar text
        $rgch            = $type;        // Built-in name type

        $unknown01       = 0x29;
        $unknown02       = 0x002b;
        $unknown03       = 0x3b;
        $unknown04       = 0xffff - $index;
        $unknown05       = 0x0000;
        $unknown06       = 0x0000;
        $unknown07       = 0x1087;
        $unknown08       = 0x8008;

        $header             = \pack('vv',  $record, $length);
        $data               = \pack('v', $grbit);
        $data              .= \pack('C', $chKey);
        $data              .= \pack('C', $cch);
        $data              .= \pack('v', $cce);
        $data              .= \pack('v', $ixals);
        $data              .= \pack('v', $itab);
        $data              .= \pack('C', $cchCustMenu);
        $data              .= \pack('C', $cchDescription);
        $data              .= \pack('C', $cchHelptopic);
        $data              .= \pack('C', $cchStatustext);
        $data              .= \pack('C', $rgch);
        $data              .= \pack('C', $unknown01);
        $data              .= \pack('v', $unknown02);
        // Column definition
        $data              .= \pack('C', $unknown03);
        $data              .= \pack('v', $unknown04);
        $data              .= \pack('v', $unknown05);
        $data              .= \pack('v', $unknown06);
        $data              .= \pack('v', $unknown07);
        $data              .= \pack('v', $unknown08);
        $data              .= \pack('v', $index);
        $data              .= \pack('v', $index);
        $data              .= \pack('v', 0x0000);
        $data              .= \pack('v', 0x3fff);
        $data              .= \pack('C', $colmin);
        $data              .= \pack('C', $colmax);
        // Row definition
        $data              .= \pack('C', $unknown03);
        $data              .= \pack('v', $unknown04);
        $data              .= \pack('v', $unknown05);
        $data              .= \pack('v', $unknown06);
        $data              .= \pack('v', $unknown07);
        $data              .= \pack('v', $unknown08);
        $data              .= \pack('v', $index);
        $data              .= \pack('v', $index);
        $data              .= \pack('v', $rowmin);
        $data              .= \pack('v', $rowmax);
        $data              .= \pack('C', 0x00);
        $data              .= \pack('C', 0xff);
        // End of data
        $data              .= \pack('C', 0x10);
        $this->_append($header . $data);
    }

    /**
     * Stores the COUNTRY record for localization.
     */
    protected function _storeCountry()
    {
        $record          = 0x008C;    // Record identifier
        $length          = 4;         // Number of bytes to follow

        $header = \pack('vv',  $record, $length);
        // using the same country code always for simplicity
        $data = \pack('vv', $this->_country_code, $this->_country_code);
        $this->_append($header . $data);
    }

    /**
     * Stores the PALETTE biff record.
     */
    protected function _storePalette()
    {
        $aref            = $this->_palette;

        $record          = 0x0092;                 // Record identifier
        $length          = 2 + 4 * \count($aref);   // Number of bytes to follow
        $ccv             =         \count($aref);   // Number of RGB values to follow
        $data = '';                                // The RGB data

        // Pack the RGB data
        foreach ($aref as $color) {
            foreach ($color as $byte) {
                $data .= \pack('C', $byte);
            }
        }

        $header = \pack('vvv',  $record, $length, $ccv);
        $this->_append($header . $data);
    }

    /**
     * Calculate
     * Handling of the SST continue blocks is complicated by the need to include an
     * additional continuation byte depending on whether the string is split between
     * blocks or whether it starts at the beginning of the block. (There are also
     * additional complications that will arise later when/if Rich Strings are
     * supported).
     */
    protected function _calculateSharedStringsSizes()
    {
        /* Iterate through the strings to calculate the CONTINUE block sizes.
           For simplicity we use the same size for the SST and CONTINUE records:
           8228 : Maximum Excel97 block size
             -4 : Length of block header
             -8 : Length of additional SST header information
             -8 : Arbitrary number to keep within _add_continue() limit = 8208
        */
        $continue_limit     = 8208;
        $block_length       = 0;
        $written            = 0;
        $this->_block_sizes = [];
        $continue           = 0;

        foreach (\array_keys($this->_str_table) as $string) {
            $string_length = \strlen($string);
            $headerinfo    = \unpack('vlength/Cencoding', $string);
            $encoding      = $headerinfo['encoding'];
            $split_string  = 0;

            // Block length is the total length of the strings that will be
            // written out in a single SST or CONTINUE block.
            $block_length += $string_length;

            // We can write the string if it doesn't cross a CONTINUE boundary
            if ($block_length < $continue_limit) {
                $written      += $string_length;

                continue;
            }

            // Deal with the cases where the next string to be written will exceed
            // the CONTINUE boundary. If the string is very long it may need to be
            // written in more than one CONTINUE record.
            while ($block_length >= $continue_limit) {
                // We need to avoid the case where a string is continued in the first
                // n bytes that contain the string header information.
                $header_length   = 3; // Min string + header size -1
                $space_remaining = $continue_limit - $written - $continue;

                /* TODO: Unicode data should only be split on char (2 byte)
                boundaries. Therefore, in some cases we need to reduce the
                amount of available
                */
                $align = 0;

                // Only applies to Unicode strings
                if (1 == $encoding) {
                    // Min string + header size -1
                    $header_length = 4;

                    if ($space_remaining > $header_length) {
                        // String contains 3 byte header => split on odd boundary
                        if (! $split_string && 1 != $space_remaining % 2) {
                            --$space_remaining;
                            $align = 1;
                        }
                        // Split section without header => split on even boundary
                        elseif ($split_string && 1 == $space_remaining % 2) {
                            --$space_remaining;
                            $align = 1;
                        }

                        $split_string = 1;
                    }
                }

                if ($space_remaining > $header_length) {
                    // Write as much as possible of the string in the current block
                    $written      += $space_remaining;

                    // Reduce the current block length by the amount written
                    $block_length -= $continue_limit - $continue - $align;

                    // Store the max size for this block
                    $this->_block_sizes[] = $continue_limit - $align;

                    // If the current string was split then the next CONTINUE block
                    // should have the string continue flag (grbit) set unless the
                    // split string fits exactly into the remaining space.
                    if ($block_length > 0) {
                        $continue = 1;
                    } else {
                        $continue = 0;
                    }
                } else {
                    // Store the max size for this block
                    $this->_block_sizes[] = $written + $continue;

                    // Not enough space to start the string in the current block
                    $block_length -= $continue_limit - $space_remaining - $continue;
                    $continue = 0;
                }

                // If the string (or substr) is small enough we can write it in the
                // new CONTINUE block. Else, go through the loop again to write it in
                // one or more CONTINUE blocks
                if ($block_length < $continue_limit) {
                    $written = $block_length;
                } else {
                    $written = 0;
                }
            }
        }

        // Store the max size for the last block unless it is empty
        if ($written + $continue) {
            $this->_block_sizes[] = $written + $continue;
        }

        /* Calculate the total length of the SST and associated CONTINUEs (if any).
         The SST record will have a length even if it contains no strings.
         This length is required to set the offsets in the BOUNDSHEET records since
         they must be written before the SST records
        */

        $tmp_block_sizes = $this->_block_sizes;

        $length  = 12;
        if (! empty($tmp_block_sizes)) {
            $length += \array_shift($tmp_block_sizes); // SST
        }
        while (! empty($tmp_block_sizes)) {
            $length += 4 + \array_shift($tmp_block_sizes); // CONTINUEs
        }

        return $length;
    }

    /**
     * Write all of the workbooks strings into an indexed array.
     * See the comments in _calculate_shared_string_sizes() for more information.
     *
     * The Excel documentation says that the SST record should be followed by an
     * EXTSST record. The EXTSST record is a hash table that is used to optimise
     * access to SST. However, despite the documentation it doesn't seem to be
     * required so we will ignore it.
     */
    protected function _storeSharedStringsTable()
    {
        $record  = 0x00fc;  // Record identifier

        // Iterate through the strings to calculate the CONTINUE block sizes
        $continue_limit = 8208;
        $block_length   = 0;
        $written        = 0;
        $continue       = 0;

        // sizes are upside down
        $tmp_block_sizes = $this->_block_sizes;
        // $tmp_block_sizes = array_reverse($this->_block_sizes);

        // The SST record is required even if it contains no strings. Thus we will
        // always have a length
        $length = 8;
        if (! empty($tmp_block_sizes)) {
            $length = 8 + \array_shift($tmp_block_sizes);
        }

        // Write the SST block header information
        $header      = \pack('vv', $record, $length);
        $data        = \pack('VV', $this->_str_total, $this->_str_unique);
        $this->_append($header . $data);

        // TODO: not good for performance
        foreach (\array_keys($this->_str_table) as $string) {
            $string_length = \strlen($string);
            $headerinfo    = \unpack('vlength/Cencoding', $string);
            $encoding      = $headerinfo['encoding'];
            $split_string  = 0;

            // Block length is the total length of the strings that will be
            // written out in a single SST or CONTINUE block.
            //
            $block_length += $string_length;

            // We can write the string if it doesn't cross a CONTINUE boundary
            if ($block_length < $continue_limit) {
                $this->_append($string);
                $written += $string_length;

                continue;
            }

            // Deal with the cases where the next string to be written will exceed
            // the CONTINUE boundary. If the string is very long it may need to be
            // written in more than one CONTINUE record.
            //
            while ($block_length >= $continue_limit) {
                // We need to avoid the case where a string is continued in the first
                // n bytes that contain the string header information.
                //
                $header_length   = 3; // Min string + header size -1
                $space_remaining = $continue_limit - $written - $continue;

                // Unicode data should only be split on char (2 byte) boundaries.
                // Therefore, in some cases we need to reduce the amount of available
                // space by 1 byte to ensure the correct alignment.
                $align = 0;

                // Only applies to Unicode strings
                if (1 == $encoding) {
                    // Min string + header size -1
                    $header_length = 4;

                    if ($space_remaining > $header_length) {
                        // String contains 3 byte header => split on odd boundary
                        if (! $split_string && 1 != $space_remaining % 2) {
                            --$space_remaining;
                            $align = 1;
                        }
                        // Split section without header => split on even boundary
                        elseif ($split_string && 1 == $space_remaining % 2) {
                            --$space_remaining;
                            $align = 1;
                        }

                        $split_string = 1;
                    }
                }

                if ($space_remaining > $header_length) {
                    // Write as much as possible of the string in the current block
                    $tmp = \substr($string, 0, $space_remaining);
                    $this->_append($tmp);

                    // The remainder will be written in the next block(s)
                    $string = \substr($string, $space_remaining);

                    // Reduce the current block length by the amount written
                    $block_length -= $continue_limit - $continue - $align;

                    // If the current string was split then the next CONTINUE block
                    // should have the string continue flag (grbit) set unless the
                    // split string fits exactly into the remaining space.
                    //
                    if ($block_length > 0) {
                        $continue = 1;
                    } else {
                        $continue = 0;
                    }
                } else {
                    // Not enough space to start the string in the current block
                    $block_length -= $continue_limit - $space_remaining - $continue;
                    $continue = 0;
                }

                // Write the CONTINUE block header
                if (! empty($this->_block_sizes)) {
                    $record  = 0x003C;
                    $length  = \array_shift($tmp_block_sizes);

                    $header  = \pack('vv', $record, $length);
                    if ($continue) {
                        $header .= \pack('C', $encoding);
                    }
                    $this->_append($header);
                }

                // If the string (or substr) is small enough we can write it in the
                // new CONTINUE block. Else, go through the loop again to write it in
                // one or more CONTINUE blocks
                //
                if ($block_length < $continue_limit) {
                    $this->_append($string);
                    $written = $block_length;
                } else {
                    $written = 0;
                }
            }
        }
    }

    /**
     * Utility function for writing formulas
     * Converts a cell's coordinates to the A1 format.
     *
     * @static
     *
     * @param int $row row for the cell to convert (0-indexed)
     * @param int $col column for the cell to convert (0-indexed)
     *
     * @return string The cell identifier in A1 format
     */
    public function rowcolToCell($row, $col)
    {
        if ($col > 255) { // maximum column value exceeded
            throw new Excel\Exception\InvalidArgumentException("Maximum column value exceeded: ${col}");
        }

        $int = (int) ($col / 26);
        $frac = $col % 26;
        $chr1 = '';

        if ($int > 0) {
            $chr1 = \chr(\ord('A') + $int - 1);
        }

        $chr2 = \chr(\ord('A') + $frac);
        ++$row;

        return $chr1 . $chr2 . $row;
    }
}
