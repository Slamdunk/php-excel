<?php

namespace Slam\Excel\Pear\Writer;

use Slam\Excel;

/**
 * Class for parsing Excel formulas.
 *
 * @author   Xavier Noguer <xnoguer@rezebra.com>
 *
 * @category FileFormats
 */
class Parser
{
    const SPREADSHEET_EXCEL_WRITER_ADD          = '+';
    const SPREADSHEET_EXCEL_WRITER_SUB          = '-';
    const SPREADSHEET_EXCEL_WRITER_MUL          = '*';
    const SPREADSHEET_EXCEL_WRITER_DIV          = '/';
    const SPREADSHEET_EXCEL_WRITER_OPEN         = '(';
    const SPREADSHEET_EXCEL_WRITER_CLOSE        = ')';
    const SPREADSHEET_EXCEL_WRITER_COMA         = ',';
    const SPREADSHEET_EXCEL_WRITER_SEMICOLON    = ';';
    const SPREADSHEET_EXCEL_WRITER_GT           = '>';
    const SPREADSHEET_EXCEL_WRITER_LT           = '<';
    const SPREADSHEET_EXCEL_WRITER_LE           = '<=';
    const SPREADSHEET_EXCEL_WRITER_GE           = '>=';
    const SPREADSHEET_EXCEL_WRITER_EQ           = '=';
    const SPREADSHEET_EXCEL_WRITER_NE           = '<>';
    const SPREADSHEET_EXCEL_WRITER_CONCAT       = '&';

    /**
     * The index of the character we are currently looking at.
     *
     * @var int
     */
    private $_current_char;

    /**
     * The token we are working on.
     *
     * @var string
     */
    private $_current_token;

    /**
     * The formula to parse.
     *
     * @var string
     */
    private $_formula;

    /**
     * The character ahead of the current char.
     *
     * @var string
     */
    private $_lookahead;

    /**
     * The parse tree to be generated.
     *
     * @var string
     */
    private $_parse_tree;

    /**
     * The byte order. 1 => big endian, 0 => little endian.
     *
     * @var int
     */
    private $_byte_order;

    /**
     * Array of external sheets.
     *
     * @var array
     */
    private $_ext_sheets;

    /**
     * Array of sheet references in the form of REF structures.
     *
     * @var array
     */
    private $_references;

    private $ptg;
    private $_functions;

    /**
     * The class constructor.
     *
     * @param int $byte_order The byte order (Little endian or Big endian) of the architecture
     */
    public function __construct($byte_order)
    {
        $this->_current_char  = 0;
        $this->_current_token = '';       // The token we are working on.
        $this->_formula       = '';       // The formula to parse.
        $this->_lookahead     = '';       // The character ahead of the current char.
        $this->_parse_tree    = '';       // The parse tree to be generated.
        $this->_initializeHashes();      // Initialize the hashes: ptg's and function's ptg's
        $this->_byte_order = $byte_order; // Little Endian or Big Endian
        $this->_ext_sheets = [];
        $this->_references = [];
    }

    /**
     * Initialize the ptg and function hashes.
     */
    private function _initializeHashes()
    {
        // The Excel ptg indices
        $this->ptg = [
            'ptgAdd'       => 0x03,
            'ptgArea'      => 0x25,
            'ptgArea3d'    => 0x3B,
            'ptgArea3dA'   => 0x7B,
            'ptgArea3dV'   => 0x5B,
            'ptgAreaA'     => 0x65,
            'ptgAreaErr'   => 0x2B,
            // 'ptgAreaErr3d' => 0x3D,
            // 'ptgAreaErr3d' => 0x5D,
            'ptgAreaErr3d' => 0x7D,
            'ptgAreaErrA'  => 0x6B,
            'ptgAreaErrV'  => 0x4B,
            'ptgAreaN'     => 0x2D,
            'ptgAreaNA'    => 0x6D,
            'ptgAreaNV'    => 0x4D,
            'ptgAreaV'     => 0x45,
            'ptgArray'     => 0x20,
            'ptgArrayA'    => 0x60,
            'ptgArrayV'    => 0x40,
            'ptgAttr'      => 0x19,
            'ptgBool'      => 0x1D,
            'ptgConcat'    => 0x08,
            'ptgDiv'       => 0x06,
            'ptgEQ'        => 0x0B,
            'ptgEndSheet'  => 0x1B,
            'ptgErr'       => 0x1C,
            'ptgExp'       => 0x01,
            'ptgFunc'      => 0x21,
            'ptgFuncA'     => 0x61,
            'ptgFuncCEA'   => 0x78,
            'ptgFuncCEV'   => 0x58,
            'ptgFuncV'     => 0x41,
            'ptgFuncVar'   => 0x22,
            'ptgFuncVarA'  => 0x62,
            'ptgFuncVarV'  => 0x42,
            'ptgGE'        => 0x0C,
            'ptgGT'        => 0x0D,
            'ptgInt'       => 0x1E,
            'ptgIsect'     => 0x0F,
            'ptgLE'        => 0x0A,
            'ptgLT'        => 0x09,
            'ptgMemArea'   => 0x26,
            'ptgMemAreaA'  => 0x66,
            'ptgMemAreaN'  => 0x2E,
            'ptgMemAreaNA' => 0x6E,
            'ptgMemAreaNV' => 0x4E,
            'ptgMemAreaV'  => 0x46,
            'ptgMemErr'    => 0x27,
            'ptgMemErrA'   => 0x67,
            'ptgMemErrV'   => 0x47,
            'ptgMemFunc'   => 0x29,
            'ptgMemFuncA'  => 0x69,
            'ptgMemFuncV'  => 0x49,
            'ptgMemNoMem'  => 0x28,
            'ptgMemNoMemA' => 0x68,
            // 'ptgMemNoMemN' => 0x2F,
            // 'ptgMemNoMemN' => 0x4F,
            'ptgMemNoMemN' => 0x6F,
            'ptgMemNoMemV' => 0x48,
            'ptgMissArg'   => 0x16,
            'ptgMul'       => 0x05,
            'ptgNE'        => 0x0E,
            'ptgName'      => 0x23,
            'ptgNameA'     => 0x63,
            'ptgNameV'     => 0x43,
            'ptgNameX'     => 0x39,
            'ptgNameXA'    => 0x79,
            'ptgNameXV'    => 0x59,
            'ptgNum'       => 0x1F,
            'ptgParen'     => 0x15,
            'ptgPercent'   => 0x14,
            'ptgPower'     => 0x07,
            'ptgRange'     => 0x11,
            'ptgRef'       => 0x24,
            'ptgRef3d'     => 0x3A,
            'ptgRef3dA'    => 0x7A,
            'ptgRef3dV'    => 0x5A,
            'ptgRefA'      => 0x64,
            'ptgRefErr'    => 0x2A,
            'ptgRefErr3d'  => 0x3C,
            'ptgRefErr3dA' => 0x7C,
            'ptgRefErr3dV' => 0x5C,
            'ptgRefErrA'   => 0x6A,
            'ptgRefErrV'   => 0x4A,
            'ptgRefN'      => 0x2C,
            'ptgRefNA'     => 0x6C,
            'ptgRefNV'     => 0x4C,
            'ptgRefV'      => 0x44,
            'ptgSheet'     => 0x1A,
            'ptgStr'       => 0x17,
            'ptgSub'       => 0x04,
            'ptgTbl'       => 0x02,
            'ptgUminus'    => 0x13,
            'ptgUnion'     => 0x10,
            'ptgUplus'     => 0x12,
        ];

        // Thanks to Michael Meeks and Gnumeric for the initial arg values.
        //
        // The following hash was generated by "function_locale.pl" in the distro.
        // Refer to function_locale.pl for non-English function names.
        //
        // The array elements are as follow:
        // ptg:   The Excel function ptg code.
        // args:  The number of arguments that the function takes:
        //           >=0 is a fixed number of arguments.
        //           -1  is a variable  number of arguments.
        // class: The reference, value or array class of the function args.
        // vol:   The function is volatile.
        //
        $this->_functions = [
              // function                  ptg  args  class  vol
              'COUNT'           => [0,   -1,    0,    0],
              'IF'              => [1,   -1,    1,    0],
              'ISNA'            => [2,    1,    1,    0],
              'ISERROR'         => [3,    1,    1,    0],
              'SUM'             => [4,   -1,    0,    0],
              'AVERAGE'         => [5,   -1,    0,    0],
              'MIN'             => [6,   -1,    0,    0],
              'MAX'             => [7,   -1,    0,    0],
              'ROW'             => [8,   -1,    0,    0],
              'COLUMN'          => [9,   -1,    0,    0],
              'NA'              => [10,    0,    0,    0],
              'NPV'             => [11,   -1,    1,    0],
              'STDEV'           => [12,   -1,    0,    0],
              'DOLLAR'          => [13,   -1,    1,    0],
              'FIXED'           => [14,   -1,    1,    0],
              'SIN'             => [15,    1,    1,    0],
              'COS'             => [16,    1,    1,    0],
              'TAN'             => [17,    1,    1,    0],
              'ATAN'            => [18,    1,    1,    0],
              'PI'              => [19,    0,    1,    0],
              'SQRT'            => [20,    1,    1,    0],
              'EXP'             => [21,    1,    1,    0],
              'LN'              => [22,    1,    1,    0],
              'LOG10'           => [23,    1,    1,    0],
              'ABS'             => [24,    1,    1,    0],
              'INT'             => [25,    1,    1,    0],
              'SIGN'            => [26,    1,    1,    0],
              'ROUND'           => [27,    2,    1,    0],
              'LOOKUP'          => [28,   -1,    0,    0],
              'INDEX'           => [29,   -1,    0,    1],
              'REPT'            => [30,    2,    1,    0],
              'MID'             => [31,    3,    1,    0],
              'LEN'             => [32,    1,    1,    0],
              'VALUE'           => [33,    1,    1,    0],
              'TRUE'            => [34,    0,    1,    0],
              'FALSE'           => [35,    0,    1,    0],
              'AND'             => [36,   -1,    0,    0],
              'OR'              => [37,   -1,    0,    0],
              'NOT'             => [38,    1,    1,    0],
              'MOD'             => [39,    2,    1,    0],
              'DCOUNT'          => [40,    3,    0,    0],
              'DSUM'            => [41,    3,    0,    0],
              'DAVERAGE'        => [42,    3,    0,    0],
              'DMIN'            => [43,    3,    0,    0],
              'DMAX'            => [44,    3,    0,    0],
              'DSTDEV'          => [45,    3,    0,    0],
              'VAR'             => [46,   -1,    0,    0],
              'DVAR'            => [47,    3,    0,    0],
              'TEXT'            => [48,    2,    1,    0],
              'LINEST'          => [49,   -1,    0,    0],
              'TREND'           => [50,   -1,    0,    0],
              'LOGEST'          => [51,   -1,    0,    0],
              'GROWTH'          => [52,   -1,    0,    0],
              'PV'              => [56,   -1,    1,    0],
              'FV'              => [57,   -1,    1,    0],
              'NPER'            => [58,   -1,    1,    0],
              'PMT'             => [59,   -1,    1,    0],
              'RATE'            => [60,   -1,    1,    0],
              'MIRR'            => [61,    3,    0,    0],
              'IRR'             => [62,   -1,    0,    0],
              'RAND'            => [63,    0,    1,    1],
              'MATCH'           => [64,   -1,    0,    0],
              'DATE'            => [65,    3,    1,    0],
              'TIME'            => [66,    3,    1,    0],
              'DAY'             => [67,    1,    1,    0],
              'MONTH'           => [68,    1,    1,    0],
              'YEAR'            => [69,    1,    1,    0],
              'WEEKDAY'         => [70,   -1,    1,    0],
              'HOUR'            => [71,    1,    1,    0],
              'MINUTE'          => [72,    1,    1,    0],
              'SECOND'          => [73,    1,    1,    0],
              'NOW'             => [74,    0,    1,    1],
              'AREAS'           => [75,    1,    0,    1],
              'ROWS'            => [76,    1,    0,    1],
              'COLUMNS'         => [77,    1,    0,    1],
              'OFFSET'          => [78,   -1,    0,    1],
              'SEARCH'          => [82,   -1,    1,    0],
              'TRANSPOSE'       => [83,    1,    1,    0],
              'TYPE'            => [86,    1,    1,    0],
              'ATAN2'           => [97,    2,    1,    0],
              'ASIN'            => [98,    1,    1,    0],
              'ACOS'            => [99,    1,    1,    0],
              'CHOOSE'          => [100,   -1,    1,    0],
              'HLOOKUP'         => [101,   -1,    0,    0],
              'VLOOKUP'         => [102,   -1,    0,    0],
              'ISREF'           => [105,    1,    0,    0],
              'LOG'             => [109,   -1,    1,    0],
              'CHAR'            => [111,    1,    1,    0],
              'LOWER'           => [112,    1,    1,    0],
              'UPPER'           => [113,    1,    1,    0],
              'PROPER'          => [114,    1,    1,    0],
              'LEFT'            => [115,   -1,    1,    0],
              'RIGHT'           => [116,   -1,    1,    0],
              'EXACT'           => [117,    2,    1,    0],
              'TRIM'            => [118,    1,    1,    0],
              'REPLACE'         => [119,    4,    1,    0],
              'SUBSTITUTE'      => [120,   -1,    1,    0],
              'CODE'            => [121,    1,    1,    0],
              'FIND'            => [124,   -1,    1,    0],
              'CELL'            => [125,   -1,    0,    1],
              'ISERR'           => [126,    1,    1,    0],
              'ISTEXT'          => [127,    1,    1,    0],
              'ISNUMBER'        => [128,    1,    1,    0],
              'ISBLANK'         => [129,    1,    1,    0],
              'T'               => [130,    1,    0,    0],
              'N'               => [131,    1,    0,    0],
              'DATEVALUE'       => [140,    1,    1,    0],
              'TIMEVALUE'       => [141,    1,    1,    0],
              'SLN'             => [142,    3,    1,    0],
              'SYD'             => [143,    4,    1,    0],
              'DDB'             => [144,   -1,    1,    0],
              'INDIRECT'        => [148,   -1,    1,    1],
              'CALL'            => [150,   -1,    1,    0],
              'CLEAN'           => [162,    1,    1,    0],
              'MDETERM'         => [163,    1,    2,    0],
              'MINVERSE'        => [164,    1,    2,    0],
              'MMULT'           => [165,    2,    2,    0],
              'IPMT'            => [167,   -1,    1,    0],
              'PPMT'            => [168,   -1,    1,    0],
              'COUNTA'          => [169,   -1,    0,    0],
              'PRODUCT'         => [183,   -1,    0,    0],
              'FACT'            => [184,    1,    1,    0],
              'DPRODUCT'        => [189,    3,    0,    0],
              'ISNONTEXT'       => [190,    1,    1,    0],
              'STDEVP'          => [193,   -1,    0,    0],
              'VARP'            => [194,   -1,    0,    0],
              'DSTDEVP'         => [195,    3,    0,    0],
              'DVARP'           => [196,    3,    0,    0],
              'TRUNC'           => [197,   -1,    1,    0],
              'ISLOGICAL'       => [198,    1,    1,    0],
              'DCOUNTA'         => [199,    3,    0,    0],
              'ROUNDUP'         => [212,    2,    1,    0],
              'ROUNDDOWN'       => [213,    2,    1,    0],
              'RANK'            => [216,   -1,    0,    0],
              'ADDRESS'         => [219,   -1,    1,    0],
              'DAYS360'         => [220,   -1,    1,    0],
              'TODAY'           => [221,    0,    1,    1],
              'VDB'             => [222,   -1,    1,    0],
              'MEDIAN'          => [227,   -1,    0,    0],
              'SUMPRODUCT'      => [228,   -1,    2,    0],
              'SINH'            => [229,    1,    1,    0],
              'COSH'            => [230,    1,    1,    0],
              'TANH'            => [231,    1,    1,    0],
              'ASINH'           => [232,    1,    1,    0],
              'ACOSH'           => [233,    1,    1,    0],
              'ATANH'           => [234,    1,    1,    0],
              'DGET'            => [235,    3,    0,    0],
              'INFO'            => [244,    1,    1,    1],
              'DB'              => [247,   -1,    1,    0],
              'FREQUENCY'       => [252,    2,    0,    0],
              'ERROR.TYPE'      => [261,    1,    1,    0],
              'REGISTER.ID'     => [267,   -1,    1,    0],
              'AVEDEV'          => [269,   -1,    0,    0],
              'BETADIST'        => [270,   -1,    1,    0],
              'GAMMALN'         => [271,    1,    1,    0],
              'BETAINV'         => [272,   -1,    1,    0],
              'BINOMDIST'       => [273,    4,    1,    0],
              'CHIDIST'         => [274,    2,    1,    0],
              'CHIINV'          => [275,    2,    1,    0],
              'COMBIN'          => [276,    2,    1,    0],
              'CONFIDENCE'      => [277,    3,    1,    0],
              'CRITBINOM'       => [278,    3,    1,    0],
              'EVEN'            => [279,    1,    1,    0],
              'EXPONDIST'       => [280,    3,    1,    0],
              'FDIST'           => [281,    3,    1,    0],
              'FINV'            => [282,    3,    1,    0],
              'FISHER'          => [283,    1,    1,    0],
              'FISHERINV'       => [284,    1,    1,    0],
              'FLOOR'           => [285,    2,    1,    0],
              'GAMMADIST'       => [286,    4,    1,    0],
              'GAMMAINV'        => [287,    3,    1,    0],
              'CEILING'         => [288,    2,    1,    0],
              'HYPGEOMDIST'     => [289,    4,    1,    0],
              'LOGNORMDIST'     => [290,    3,    1,    0],
              'LOGINV'          => [291,    3,    1,    0],
              'NEGBINOMDIST'    => [292,    3,    1,    0],
              'NORMDIST'        => [293,    4,    1,    0],
              'NORMSDIST'       => [294,    1,    1,    0],
              'NORMINV'         => [295,    3,    1,    0],
              'NORMSINV'        => [296,    1,    1,    0],
              'STANDARDIZE'     => [297,    3,    1,    0],
              'ODD'             => [298,    1,    1,    0],
              'PERMUT'          => [299,    2,    1,    0],
              'POISSON'         => [300,    3,    1,    0],
              'TDIST'           => [301,    3,    1,    0],
              'WEIBULL'         => [302,    4,    1,    0],
              'SUMXMY2'         => [303,    2,    2,    0],
              'SUMX2MY2'        => [304,    2,    2,    0],
              'SUMX2PY2'        => [305,    2,    2,    0],
              'CHITEST'         => [306,    2,    2,    0],
              'CORREL'          => [307,    2,    2,    0],
              'COVAR'           => [308,    2,    2,    0],
              'FORECAST'        => [309,    3,    2,    0],
              'FTEST'           => [310,    2,    2,    0],
              'INTERCEPT'       => [311,    2,    2,    0],
              'Excel_PEARSON'         => [312,    2,    2,    0],
              'RSQ'             => [313,    2,    2,    0],
              'STEYX'           => [314,    2,    2,    0],
              'SLOPE'           => [315,    2,    2,    0],
              'TTEST'           => [316,    4,    2,    0],
              'PROB'            => [317,   -1,    2,    0],
              'DEVSQ'           => [318,   -1,    0,    0],
              'GEOMEAN'         => [319,   -1,    0,    0],
              'HARMEAN'         => [320,   -1,    0,    0],
              'SUMSQ'           => [321,   -1,    0,    0],
              'KURT'            => [322,   -1,    0,    0],
              'SKEW'            => [323,   -1,    0,    0],
              'ZTEST'           => [324,   -1,    0,    0],
              'LARGE'           => [325,    2,    0,    0],
              'SMALL'           => [326,    2,    0,    0],
              'QUARTILE'        => [327,    2,    0,    0],
              'PERCENTILE'      => [328,    2,    0,    0],
              'PERCENTRANK'     => [329,   -1,    0,    0],
              'MODE'            => [330,   -1,    2,    0],
              'TRIMMEAN'        => [331,    2,    0,    0],
              'TINV'            => [332,    2,    1,    0],
              'CONCATENATE'     => [336,   -1,    1,    0],
              'POWER'           => [337,    2,    1,    0],
              'RADIANS'         => [342,    1,    1,    0],
              'DEGREES'         => [343,    1,    1,    0],
              'SUBTOTAL'        => [344,   -1,    0,    0],
              'SUMIF'           => [345,   -1,    0,    0],
              'COUNTIF'         => [346,    2,    0,    0],
              'COUNTBLANK'      => [347,    1,    0,    0],
              'ROMAN'           => [354,   -1,    1,    0],
              ];
    }

    /**
     * Convert a token to the proper ptg value.
     *
     *
     * @param mixed $token the token to convert
     *
     * @return mixed the converted token on success. Excel_PEAR_Error if the token
     *               is not recognized
     */
    private function _convert($token)
    {
        if (\preg_match('/^"[^"]{0,255}"$/', $token)) {
            return $this->_convertString($token);
        }
        if (\is_numeric($token)) {
            return $this->_convertNumber($token);
        // match references like A1 or $A$1
        }
        if (\preg_match('/^\$?([A-Ia-i]?[A-Za-z])\$?(\d+)$/', $token)) {
            return $this->_convertRef2d($token);
        // match external references like Sheet1!A1 or Sheet1:Sheet2!A1
        }
        if (\preg_match('/^\\w+(\\:\\w+)?\\![A-Ia-i]?[A-Za-z](\\d+)$/u', $token)) {
            return $this->_convertRef3d($token);
        // match external references like 'Sheet1'!A1 or 'Sheet1:Sheet2'!A1
        }
        if (\preg_match("/^'[\\w -]+(\\:[\\w -]+)?'\\![A-Ia-i]?[A-Za-z](\\d+)$/u", $token)) {
            return $this->_convertRef3d($token);
        // match ranges like A1:B2
        }
        if (\preg_match('/^($)?[A-Ia-i]?[A-Za-z]($)?(\\d+)\\:($)?[A-Ia-i]?[A-Za-z]($)?(\\d+)$/', $token)) {
            return $this->_convertRange2d($token);
        // match ranges like A1..B2
        }
        if (\preg_match('/^($)?[A-Ia-i]?[A-Za-z]($)?(\\d+)\\.\\.($)?[A-Ia-i]?[A-Za-z]($)?(\\d+)$/', $token)) {
            return $this->_convertRange2d($token);
        // match external ranges like Sheet1!A1 or Sheet1:Sheet2!A1:B2
        }
        if (\preg_match('/^\\w+(\\:\\w+)?\\!([A-Ia-i]?[A-Za-z])?(\\d+)\\:([A-Ia-i]?[A-Za-z])?(\\d+)$/u', $token)) {
            return $this->_convertRange3d($token);
        // match external ranges like 'Sheet1'!A1 or 'Sheet1:Sheet2'!A1:B2
        }
        if (\preg_match("/^'[\\w -]+(\\:[\\w -]+)?'\\!([A-Ia-i]?[A-Za-z])?(\\d+)\\:([A-Ia-i]?[A-Za-z])?(\\d+)$/u", $token)) {
            return $this->_convertRange3d($token);
        // operators (including parentheses)
        }
        if (isset($this->ptg[$token])) {
            return \pack('C', $this->ptg[$token]);
        // commented so argument number can be processed correctly. See toReversePolish().
        /*elseif (preg_match("/[A-Z0-9\xc0-\xdc\.]+/",$token))
        {
            return($this->_convertFunction($token,$this->_func_args));
        }*/

        // if it's an argument, ignore the token (the argument remains)
        }
        if ('arg' == $token) {
            return '';
        }
        // TODO: use real error codes
        throw new Excel\Exception\RuntimeException("Unknown token ${token}");
    }

    /**
     * Convert a number token to ptgInt or ptgNum.
     *
     *
     * @param mixed $num an integer or double for conversion to its ptg value
     */
    private function _convertNumber($num)
    {
        // Integer in the range 0..2**16-1
        if ((\preg_match('/^\\d+$/', $num)) and ($num <= 65535)) {
            return \pack('Cv', $this->ptg['ptgInt'], $num);
        }   // A float
            if ($this->_byte_order) { // if it's Big Endian
                $num = \strrev($num);
            }

        return \pack('Cd', $this->ptg['ptgNum'], $num);
    }

    /**
     * Convert a string token to ptgStr.
     *
     *
     * @param string $string a string for conversion to its ptg value
     *
     * @return mixed the converted token on success. Excel_PEAR_Error if the string
     *               is longer than 255 characters.
     */
    private function _convertString($string)
    {
        // chop away beggining and ending quotes
        $string = \substr($string, 1, \strlen($string) - 2);
        if (\strlen($string) > 255) {
            throw new Excel\Exception\RuntimeException('String is too long');
        }

        return \pack('CC', $this->ptg['ptgStr'], \strlen($string)) . $string;
    }

    /**
     * Convert a function to a ptgFunc or ptgFuncVarV depending on the number of
     * args that it takes.
     *
     *
     * @param string $token    the name of the function for convertion to ptg value
     * @param int    $num_args the number of arguments the function receives
     *
     * @return string The packed ptg for the function
     */
    private function _convertFunction($token, $num_args)
    {
        $args     = $this->_functions[$token][1];

        // Fixed number of args eg. TIME($i,$j,$k).
        if ($args >= 0) {
            return \pack('Cv', $this->ptg['ptgFuncV'], $this->_functions[$token][0]);
        }
        // Variable number of args eg. SUM($i,$j,$k, ..).
        if ($args == -1) {
            return \pack('CCv', $this->ptg['ptgFuncVarV'], $num_args, $this->_functions[$token][0]);
        }
    }

    /**
     * Convert an Excel range such as A1:D4 to a ptgRefV.
     *
     *
     * @param string $range An Excel range in the A1:A2 or A1..A2 format.
     */
    private function _convertRange2d($range, $class = 0)
    {
        // TODO: possible class value 0,1,2 check Formula.pm
        // Split the range into 2 cell refs
        if (\preg_match('/^([A-Ia-i]?[A-Za-z])(\\d+)\\:([A-Ia-i]?[A-Za-z])(\\d+)$/', $range)) {
            list($cell1, $cell2) = \explode(':', $range);
        } elseif (\preg_match('/^([A-Ia-i]?[A-Za-z])(\\d+)\\.\\.([A-Ia-i]?[A-Za-z])(\\d+)$/', $range)) {
            list($cell1, $cell2) = \explode('..', $range);
        } else {
            // TODO: use real error codes
            throw new Excel\Exception\RuntimeException('Unknown range separator');
        }

        // Convert the cell references
        $cell_array1 = $this->_cellToPackedRowcol($cell1);
        list($row1, $col1) = $cell_array1;
        $cell_array2 = $this->_cellToPackedRowcol($cell2);
        list($row2, $col2) = $cell_array2;

        // The ptg value depends on the class of the ptg.
        if (0 == $class) {
            $ptgArea = \pack('C', $this->ptg['ptgArea']);
        } elseif (1 == $class) {
            $ptgArea = \pack('C', $this->ptg['ptgAreaV']);
        } elseif (2 == $class) {
            $ptgArea = \pack('C', $this->ptg['ptgAreaA']);
        } else {
            // TODO: use real error codes
            throw new Excel\Exception\RuntimeException("Unknown class ${class}");
        }

        return $ptgArea . $row1 . $row2 . $col1 . $col2;
    }

    /**
     * Convert an Excel 3d range such as "Sheet1!A1:D4" or "Sheet1:Sheet2!A1:D4" to
     * a ptgArea3d.
     *
     *
     * @param string $token an Excel range in the Sheet1!A1:A2 format
     *
     * @return mixed the packed ptgArea3d token on success, Excel_PEAR_Error on failure
     */
    private function _convertRange3d($token)
    {
        $class = 2; // as far as I know, this is magick.

        // Split the ref at the ! symbol
        list($ext_ref, $range) = \explode('!', $token);

        // Convert the external reference part (different for BIFF8)
        $ext_ref = $this->_packExtRef($ext_ref);

        // Split the range into 2 cell refs
        list($cell1, $cell2) = \explode(':', $range);

        // Convert the cell references
        if (\preg_match('/^($)?[A-Ia-i]?[A-Za-z]($)?(\\d+)$/', $cell1)) {
            $cell_array1 = $this->_cellToPackedRowcol($cell1);
            list($row1, $col1) = $cell_array1;
            $cell_array2 = $this->_cellToPackedRowcol($cell2);
            list($row2, $col2) = $cell_array2;
        } else { // It's a rows range (like 26:27)
            $cells_array = $this->_rangeToPackedRange($cell1 . ':' . $cell2);
            list($row1, $col1, $row2, $col2) = $cells_array;
        }

        // The ptg value depends on the class of the ptg.
        if (0 == $class) {
            $ptgArea = \pack('C', $this->ptg['ptgArea3d']);
        } elseif (1 == $class) {
            $ptgArea = \pack('C', $this->ptg['ptgArea3dV']);
        } elseif (2 == $class) {
            $ptgArea = \pack('C', $this->ptg['ptgArea3dA']);
        } else {
            throw new Excel\Exception\RuntimeException("Unknown class ${class}");
        }

        return $ptgArea . $ext_ref . $row1 . $row2 . $col1 . $col2;
    }

    /**
     * Convert an Excel reference such as A1, $B2, C$3 or $D$4 to a ptgRefV.
     *
     *
     * @param string $cell An Excel cell reference
     *
     * @return string The cell in packed() format with the corresponding ptg
     */
    private function _convertRef2d($cell)
    {
        $class = 2; // as far as I know, this is magick.

        // Convert the cell reference
        $cell_array = $this->_cellToPackedRowcol($cell);
        list($row, $col) = $cell_array;

        // The ptg value depends on the class of the ptg.
        if (0 == $class) {
            $ptgRef = \pack('C', $this->ptg['ptgRef']);
        } elseif (1 == $class) {
            $ptgRef = \pack('C', $this->ptg['ptgRefV']);
        } elseif (2 == $class) {
            $ptgRef = \pack('C', $this->ptg['ptgRefA']);
        } else {
            // TODO: use real error codes
            throw new Excel\Exception\RuntimeException("Unknown class ${class}");
        }

        return $ptgRef . $row . $col;
    }

    /**
     * Convert an Excel 3d reference such as "Sheet1!A1" or "Sheet1:Sheet2!A1" to a
     * ptgRef3d.
     *
     *
     * @param string $cell An Excel cell reference
     *
     * @return mixed the packed ptgRef3d token on success, Excel_PEAR_Error on failure
     */
    private function _convertRef3d($cell)
    {
        $class = 2; // as far as I know, this is magick.

        // Split the ref at the ! symbol
        list($ext_ref, $cell) = \explode('!', $cell);

        // Convert the external reference part (different for BIFF8)
        $ext_ref = $this->_packExtRef($ext_ref);

        // Convert the cell reference part
        list($row, $col) = $this->_cellToPackedRowcol($cell);

        // The ptg value depends on the class of the ptg.
        if (0 == $class) {
            $ptgRef = \pack('C', $this->ptg['ptgRef3d']);
        } elseif (1 == $class) {
            $ptgRef = \pack('C', $this->ptg['ptgRef3dV']);
        } elseif (2 == $class) {
            $ptgRef = \pack('C', $this->ptg['ptgRef3dA']);
        } else {
            throw new Excel\Exception\RuntimeException("Unknown class ${class}");
        }

        return $ptgRef . $ext_ref . $row . $col;
    }

    /**
     * Convert the sheet name part of an external reference, for example "Sheet1" or
     * "Sheet1:Sheet2", to a packed structure.
     *
     *
     * @param string $ext_ref The name of the external reference
     *
     * @return string The reference index in packed() format
     */
    private function _packExtRef($ext_ref)
    {
        $ext_ref = \preg_replace("/^'/", '', $ext_ref); // Remove leading  ' if any.
        $ext_ref = \preg_replace("/'$/", '', $ext_ref); // Remove trailing ' if any.

        // Check if there is a sheet range eg., Sheet1:Sheet2.
        if (\preg_match('/:/', $ext_ref)) {
            list($sheet_name1, $sheet_name2) = \explode(':', $ext_ref);

            $sheet1 = $this->_getSheetIndex($sheet_name1);
            if ($sheet1 == -1) {
                throw new Excel\Exception\RuntimeException("Unknown sheet name ${sheet_name1} in formula");
            }
            $sheet2 = $this->_getSheetIndex($sheet_name2);
            if ($sheet2 == -1) {
                throw new Excel\Exception\RuntimeException("Unknown sheet name ${sheet_name2} in formula");
            }

            // Reverse max and min sheet numbers if necessary
            if ($sheet1 > $sheet2) {
                list($sheet1, $sheet2) = [$sheet2, $sheet1];
            }
        } else { // Single sheet name only.
            $sheet1 = $this->_getSheetIndex($ext_ref);
            if ($sheet1 == -1) {
                throw new Excel\Exception\RuntimeException("Unknown sheet name ${ext_ref} in formula");
            }
            $sheet2 = $sheet1;
        }

        // References are stored relative to 0xFFFF.
        $offset = -1 - $sheet1;

        return \pack('vdvv', $offset, 0x00, $sheet1, $sheet2);
    }

    /**
     * Look up the REF index that corresponds to an external sheet name
     * (or range). If it doesn't exist yet add it to the workbook's references
     * array. It assumes all sheet names given must exist.
     *
     *
     * @param string $ext_ref The name of the external reference
     *
     * @return mixed The reference index in packed() format on success,
     *               Excel_PEAR_Error on failure
     */
    /*
    private function _getRefIndex($ext_ref)
    {
        $ext_ref = preg_replace("/^'/", '', $ext_ref); // Remove leading  ' if any.
        $ext_ref = preg_replace("/'$/", '', $ext_ref); // Remove trailing ' if any.

        // Check if there is a sheet range eg., Sheet1:Sheet2.
        if (preg_match('/:/', $ext_ref)) {
            list($sheet_name1, $sheet_name2) = explode(':', $ext_ref);

            $sheet1 = $this->_getSheetIndex($sheet_name1);
            if ($sheet1 == -1) {
                throw new Excel\Exception\RuntimeException("Unknown sheet name $sheet_name1 in formula");
            }
            $sheet2 = $this->_getSheetIndex($sheet_name2);
            if ($sheet2 == -1) {
                throw new Excel\Exception\RuntimeException("Unknown sheet name $sheet_name2 in formula");
            }

            // Reverse max and min sheet numbers if necessary
            if ($sheet1 > $sheet2) {
                list($sheet1, $sheet2) = array($sheet2, $sheet1);
            }
        } else { // Single sheet name only.
            $sheet1 = $this->_getSheetIndex($ext_ref);
            if ($sheet1 == -1) {
                throw new Excel\Exception\RuntimeException("Unknown sheet name $ext_ref in formula");
            }
            $sheet2 = $sheet1;
        }

        // assume all references belong to this document
        $supbook_index = 0x00;
        $ref = pack('vvv', $supbook_index, $sheet1, $sheet2);
        $total_references = count($this->_references);
        $index = -1;
        for ($i = 0; $i < $total_references; ++$i) {
            if ($ref == $this->_references[$i]) {
                $index = $i;
                break;
            }
        }
        // if REF was not found add it to references array
        if ($index == -1) {
            $this->_references[$total_references] = $ref;
            $index = $total_references;
        }

        return pack('v', $index);
    }
    */

    /**
     * Look up the index that corresponds to an external sheet name. The hash of
     * sheet names is updated by the addworksheet() method of the
     * Excel_Writer_Workbook class.
     *
     *
     * @return int The sheet index, -1 if the sheet was not found
     */
    private function _getSheetIndex($sheet_name)
    {
        if (! isset($this->_ext_sheets[$sheet_name])) {
            return -1;
        }

        return $this->_ext_sheets[$sheet_name];
    }

    /**
     * This method is used to update the array of sheet names. It is
     * called by the addWorksheet() method of the
     * Excel_Writer_Workbook class.
     *
     *
     * @see Excel_Writer_Workbook::addWorksheet()
     *
     * @param string $name  The name of the worksheet being added
     * @param int    $index The index of the worksheet being added
     */
    public function setExtSheet($name, $index)
    {
        $this->_ext_sheets[$name] = $index;
    }

    /**
     * pack() row and column into the required 3 or 4 byte format.
     *
     *
     * @param string $cell The Excel cell reference to be packed
     *
     * @return array Array containing the row and column in packed() format
     */
    private function _cellToPackedRowcol($cell)
    {
        $cell = \strtoupper($cell);
        list($row, $col, $row_rel, $col_rel) = $this->_cellToRowcol($cell);
        if ($col >= 256) {
            throw new Excel\Exception\RuntimeException("Column in: ${cell} greater than 255");
        }
        if ($row >= 65536) {
            throw new Excel\Exception\RuntimeException("Row in: ${cell} greater than 65536");
        }

        // Set the high bits to indicate if row or col are relative.
        $row    |= $col_rel << 14;
        $row    |= $row_rel << 15;
        $col     = \pack('C', $col);

        $row     = \pack('v', $row);

        return [$row, $col];
    }

    /**
     * pack() row range into the required 3 or 4 byte format.
     * Just using maximum col/rows, which is probably not the correct solution.
     *
     *
     * @param string $range The Excel range to be packed
     *
     * @return array Array containing (row1,col1,row2,col2) in packed() format
     */
    private function _rangeToPackedRange($range)
    {
        \preg_match('/(\$)?(\d+)\:(\$)?(\d+)/', $range, $match);
        // return absolute rows if there is a $ in the ref
        $row1_rel = empty($match[1]) ? 1 : 0;
        $row1     = $match[2];
        $row2_rel = empty($match[3]) ? 1 : 0;
        $row2     = $match[4];
        // Convert 1-index to zero-index
        --$row1;
        --$row2;
        // Trick poor innocent Excel
        $col1 = 0;
        $col2 = 16383; // FIXME: maximum possible value for Excel 5 (change this!!!)

        if (($row1 >= 65536) or ($row2 >= 65536)) {
            throw new Excel\Exception\RuntimeException("Row in: ${range} greater than 65536");
        }

        // Set the high bits to indicate if rows are relative.
        $row1    |= $row1_rel << 14; // FIXME: probably a bug
        $row2    |= $row2_rel << 15;
        $col1     = \pack('C', $col1);
        $col2     = \pack('C', $col2);
        $row1     = \pack('v', $row1);
        $row2     = \pack('v', $row2);

        return [$row1, $col1, $row2, $col2];
    }

    /**
     * Convert an Excel cell reference such as A1 or $B2 or C$3 or $D$4 to a zero
     * indexed row and column number. Also returns two (0,1) values to indicate
     * whether the row or column are relative references.
     *
     *
     * @param string $cell the Excel cell reference in A1 format
     *
     * @return array
     */
    private function _cellToRowcol($cell)
    {
        \preg_match('/(\$)?([A-I]?[A-Z])(\$)?(\d+)/', $cell, $match);
        // return absolute column if there is a $ in the ref
        $col_rel = empty($match[1]) ? 1 : 0;
        $col_ref = $match[2];
        $row_rel = empty($match[3]) ? 1 : 0;
        $row     = $match[4];

        // Convert base26 column string to a number.
        $expn   = \strlen($col_ref) - 1;
        $col    = 0;
        $col_ref_length = \strlen($col_ref);
        for ($i = 0; $i < $col_ref_length; ++$i) {
            $col += (\ord($col_ref[$i]) - \ord('A') + 1) * \pow(26, $expn);
            --$expn;
        }

        // Convert 1-index to zero-index
        --$row;
        --$col;

        return [$row, $col, $row_rel, $col_rel];
    }

    /**
     * Advance to the next valid token.
     */
    private function _advance()
    {
        $i = $this->_current_char;
        $formula_length = \strlen($this->_formula);
        // eat up white spaces
        if ($i < $formula_length) {
            while (' ' == $this->_formula[$i]) {
                ++$i;
            }

            if ($i < ($formula_length - 1)) {
                $this->_lookahead = $this->_formula[$i + 1];
            }
        }

        $token = '';
        while ($i < $formula_length) {
            $token .= $this->_formula[$i];
            if ($i < ($formula_length - 1)) {
                $this->_lookahead = $this->_formula[$i + 1];
            } else {
                $this->_lookahead = '';
            }

            if ('' != $this->_match($token)) {
                // if ($i < strlen($this->_formula) - 1) {
                //    $this->_lookahead = $this->_formula{$i+1};
                // }
                $this->_current_char = $i + 1;
                $this->_current_token = $token;

                return 1;
            }

            if ($i < ($formula_length - 2)) {
                $this->_lookahead = $this->_formula[$i + 2];
            } else { // if we run out of characters _lookahead becomes empty
                $this->_lookahead = '';
            }
            ++$i;
        }
        // die("Lexical error ".$this->_current_char);
    }

    /**
     * Checks if it's a valid token.
     *
     *
     * @param mixed $token the token to check
     *
     * @return mixed The checked token or false on failure
     */
    private function _match($token)
    {
        switch ($token) {
            case self::SPREADSHEET_EXCEL_WRITER_ADD:
            case self::SPREADSHEET_EXCEL_WRITER_SUB:
            case self::SPREADSHEET_EXCEL_WRITER_MUL:
            case self::SPREADSHEET_EXCEL_WRITER_DIV:
            case self::SPREADSHEET_EXCEL_WRITER_OPEN:
            case self::SPREADSHEET_EXCEL_WRITER_CLOSE:
            case self::SPREADSHEET_EXCEL_WRITER_COMA:
            case self::SPREADSHEET_EXCEL_WRITER_GE:
            case self::SPREADSHEET_EXCEL_WRITER_LE:
            case self::SPREADSHEET_EXCEL_WRITER_EQ:
            case self::SPREADSHEET_EXCEL_WRITER_NE:
            case self::SPREADSHEET_EXCEL_WRITER_CONCAT:
            case self::SPREADSHEET_EXCEL_WRITER_SEMICOLON:
                return $token;

                break;
            case self::SPREADSHEET_EXCEL_WRITER_GT:
                if ('=' == $this->_lookahead) { // it's a GE token
                    break;
                }

                return $token;

                break;
            case self::SPREADSHEET_EXCEL_WRITER_LT:
                // it's a LE or a NE token
                if (('=' == $this->_lookahead) or ('>' == $this->_lookahead)) {
                    break;
                }

                return $token;

                break;
            default:
                // if it's a reference
                if (\preg_match('/^\$?[A-Ia-i]?[A-Za-z]\$?[0-9]+$/', $token) and
                   ! \preg_match('/[0-9]/', $this->_lookahead) and
                   (':' != $this->_lookahead) and ('.' != $this->_lookahead) and
                   ('!' != $this->_lookahead)) {
                    return $token;
                }
                // If it's an external reference (Sheet1!A1 or Sheet1:Sheet2!A1)
                if (\preg_match('/^\\w+(\\:\\w+)?\\![A-Ia-i]?[A-Za-z][0-9]+$/u', $token) and
                       ! \preg_match('/[0-9]/', $this->_lookahead) and
                       (':' != $this->_lookahead) and ('.' != $this->_lookahead)) {
                    return $token;
                }
                // If it's an external reference ('Sheet1'!A1 or 'Sheet1:Sheet2'!A1)
                if (\preg_match("/^'[\\w -]+(\\:[\\w -]+)?'\\![A-Ia-i]?[A-Za-z][0-9]+$/u", $token) and
                       ! \preg_match('/[0-9]/', $this->_lookahead) and
                       (':' != $this->_lookahead) and ('.' != $this->_lookahead)) {
                    return $token;
                }
                // if it's a range (A1:A2)
                if (\preg_match('/^($)?[A-Ia-i]?[A-Za-z]($)?[0-9]+:($)?[A-Ia-i]?[A-Za-z]($)?[0-9]+$/', $token) and
                       ! \preg_match('/[0-9]/', $this->_lookahead)) {
                    return $token;
                }
                // if it's a range (A1..A2)
                if (\preg_match('/^($)?[A-Ia-i]?[A-Za-z]($)?[0-9]+\\.\\.($)?[A-Ia-i]?[A-Za-z]($)?[0-9]+$/', $token) and
                       ! \preg_match('/[0-9]/', $this->_lookahead)) {
                    return $token;
                }
                // If it's an external range like Sheet1!A1 or Sheet1:Sheet2!A1:B2
                if (\preg_match('/^\\w+(\\:\\w+)?\\!([A-Ia-i]?[A-Za-z])?[0-9]+:([A-Ia-i]?[A-Za-z])?[0-9]+$/u', $token) and
                       ! \preg_match('/[0-9]/', $this->_lookahead)) {
                    return $token;
                }
                // If it's an external range like 'Sheet1'!A1 or 'Sheet1:Sheet2'!A1:B2
                if (\preg_match("/^'[\\w -]+(\\:[\\w -]+)?'\\!([A-Ia-i]?[A-Za-z])?[0-9]+:([A-Ia-i]?[A-Za-z])?[0-9]+$/u", $token) and
                       ! \preg_match('/[0-9]/', $this->_lookahead)) {
                    return $token;
                }
                // If it's a number (check that it's not a sheet name or range)
                if (\is_numeric($token) and
                        (! \is_numeric($token . $this->_lookahead) or ('' == $this->_lookahead)) and
                        ('!' != $this->_lookahead) and (':' != $this->_lookahead)) {
                    return $token;
                }
                // If it's a string (of maximum 255 characters)
                if (\preg_match('/^"[^"]{0,255}"$/', $token)) {
                    return $token;
                }
                // if it's a function call
                if (\preg_match("/^[A-Z0-9\xc0-\xdc\\.]+$/i", $token) and ('(' == $this->_lookahead)) {
                    return $token;
                }

                return '';
        }
    }

    /**
     * The parsing method. It parses a formula.
     *
     *
     * @param string $formula the formula to parse, without the initial equal
     *                        sign (=)
     *
     * @return mixed true on success, Excel_PEAR_Error on failure
     */
    public function parse($formula)
    {
        $this->_current_char = 0;
        $this->_formula      = $formula;
        $this->_lookahead    = $formula[1];
        $this->_advance();
        $this->_parse_tree   = $this->_condition();

        return true;
    }

    /**
     * It parses a condition. It assumes the following rule:
     * Cond -> Expr [(">" | "<") Expr].
     *
     *
     * @return mixed The parsed ptg'd tree on success, Excel_PEAR_Error on failure
     */
    private function _condition()
    {
        $result = $this->_expression();
        if (self::SPREADSHEET_EXCEL_WRITER_LT == $this->_current_token) {
            $this->_advance();
            $result2 = $this->_expression();
            $result = $this->_createTree('ptgLT', $result, $result2);
        } elseif (self::SPREADSHEET_EXCEL_WRITER_GT == $this->_current_token) {
            $this->_advance();
            $result2 = $this->_expression();
            $result = $this->_createTree('ptgGT', $result, $result2);
        } elseif (self::SPREADSHEET_EXCEL_WRITER_LE == $this->_current_token) {
            $this->_advance();
            $result2 = $this->_expression();
            $result = $this->_createTree('ptgLE', $result, $result2);
        } elseif (self::SPREADSHEET_EXCEL_WRITER_GE == $this->_current_token) {
            $this->_advance();
            $result2 = $this->_expression();
            $result = $this->_createTree('ptgGE', $result, $result2);
        } elseif (self::SPREADSHEET_EXCEL_WRITER_EQ == $this->_current_token) {
            $this->_advance();
            $result2 = $this->_expression();
            $result = $this->_createTree('ptgEQ', $result, $result2);
        } elseif (self::SPREADSHEET_EXCEL_WRITER_NE == $this->_current_token) {
            $this->_advance();
            $result2 = $this->_expression();
            $result = $this->_createTree('ptgNE', $result, $result2);
        } elseif (self::SPREADSHEET_EXCEL_WRITER_CONCAT == $this->_current_token) {
            $this->_advance();
            $result2 = $this->_expression();
            $result = $this->_createTree('ptgConcat', $result, $result2);
        }

        return $result;
    }

    /**
     * It parses a expression. It assumes the following rule:
     * Expr -> Term [("+" | "-") Term]
     *      -> "string"
     *      -> "-" Term.
     *
     *
     * @return mixed The parsed ptg'd tree on success, Excel_PEAR_Error on failure
     */
    private function _expression()
    {
        // If it's a string return a string node
        if (\preg_match('/^"[^"]{0,255}"$/', $this->_current_token)) {
            $result = $this->_createTree($this->_current_token, '', '');
            $this->_advance();

            return $result;
        }
        if (self::SPREADSHEET_EXCEL_WRITER_SUB == $this->_current_token) {
            // catch "-" Term
            $this->_advance();
            $result2 = $this->_expression();
            $result = $this->_createTree('ptgUminus', $result2, '');

            return $result;
        }
        $result = $this->_term();
        while ((self::SPREADSHEET_EXCEL_WRITER_ADD == $this->_current_token) or
               (self::SPREADSHEET_EXCEL_WRITER_SUB == $this->_current_token)) {
            if (self::SPREADSHEET_EXCEL_WRITER_ADD == $this->_current_token) {
                $this->_advance();
                $result2 = $this->_term();
                $result = $this->_createTree('ptgAdd', $result, $result2);
            } else {
                $this->_advance();
                $result2 = $this->_term();
                $result = $this->_createTree('ptgSub', $result, $result2);
            }
        }

        return $result;
    }

    /**
     * This function just introduces a ptgParen element in the tree, so that Excel
     * doesn't get confused when working with a parenthesized formula afterwards.
     *
     *
     * @see _fact()
     *
     * @return array The parsed ptg'd tree
     */
    private function _parenthesizedExpression()
    {
        $result = $this->_createTree('ptgParen', $this->_expression(), '');

        return $result;
    }

    /**
     * It parses a term. It assumes the following rule:
     * Term -> Fact [("*" | "/") Fact].
     *
     *
     * @return mixed The parsed ptg'd tree on success, Excel_PEAR_Error on failure
     */
    private function _term()
    {
        $result = $this->_fact();
        while ((self::SPREADSHEET_EXCEL_WRITER_MUL == $this->_current_token) or
               (self::SPREADSHEET_EXCEL_WRITER_DIV == $this->_current_token)) {
            if (self::SPREADSHEET_EXCEL_WRITER_MUL == $this->_current_token) {
                $this->_advance();
                $result2 = $this->_fact();
                $result = $this->_createTree('ptgMul', $result, $result2);
            } else {
                $this->_advance();
                $result2 = $this->_fact();
                $result = $this->_createTree('ptgDiv', $result, $result2);
            }
        }

        return $result;
    }

    /**
     * It parses a factor. It assumes the following rule:
     * Fact -> ( Expr )
     *       | CellRef
     *       | CellRange
     *       | Number
     *       | Function.
     *
     *
     * @return mixed The parsed ptg'd tree on success, Excel_PEAR_Error on failure
     */
    private function _fact()
    {
        if (self::SPREADSHEET_EXCEL_WRITER_OPEN == $this->_current_token) {
            $this->_advance();         // eat the "("
            $result = $this->_parenthesizedExpression();
            if (self::SPREADSHEET_EXCEL_WRITER_CLOSE != $this->_current_token) {
                throw new Excel\Exception\RuntimeException("')' token expected.");
            }
            $this->_advance();         // eat the ")"
            return $result;
        }
        // if it's a reference
        if (\preg_match('/^\$?[A-Ia-i]?[A-Za-z]\$?[0-9]+$/', $this->_current_token)) {
            $result = $this->_createTree($this->_current_token, '', '');
            $this->_advance();

            return $result;
        }
        // If it's an external reference (Sheet1!A1 or Sheet1:Sheet2!A1)
        if (\preg_match('/^\\w+(\\:\\w+)?\\![A-Ia-i]?[A-Za-z][0-9]+$/u', $this->_current_token)) {
            $result = $this->_createTree($this->_current_token, '', '');
            $this->_advance();

            return $result;
        }
        // If it's an external reference ('Sheet1'!A1 or 'Sheet1:Sheet2'!A1)
        if (\preg_match("/^'[\\w -]+(\\:[\\w -]+)?'\\![A-Ia-i]?[A-Za-z][0-9]+$/u", $this->_current_token)) {
            $result = $this->_createTree($this->_current_token, '', '');
            $this->_advance();

            return $result;
        }
        // if it's a range
        if (\preg_match('/^($)?[A-Ia-i]?[A-Za-z]($)?[0-9]+:($)?[A-Ia-i]?[A-Za-z]($)?[0-9]+$/', $this->_current_token) or
                \preg_match('/^($)?[A-Ia-i]?[A-Za-z]($)?[0-9]+\\.\\.($)?[A-Ia-i]?[A-Za-z]($)?[0-9]+$/', $this->_current_token)) {
            $result = $this->_current_token;
            $this->_advance();

            return $result;
        }
        // If it's an external range (Sheet1!A1 or Sheet1!A1:B2)
        if (\preg_match('/^\\w+(\\:\\w+)?\\!([A-Ia-i]?[A-Za-z])?[0-9]+:([A-Ia-i]?[A-Za-z])?[0-9]+$/u', $this->_current_token)) {
            $result = $this->_current_token;
            $this->_advance();

            return $result;
        }
        // If it's an external range ('Sheet1'!A1 or 'Sheet1'!A1:B2)
        if (\preg_match("/^'[\\w -]+(\\:[\\w -]+)?'\\!([A-Ia-i]?[A-Za-z])?[0-9]+:([A-Ia-i]?[A-Za-z])?[0-9]+$/u", $this->_current_token)) {
            $result = $this->_current_token;
            $this->_advance();

            return $result;
        }
        if (\is_numeric($this->_current_token)) {
            $result = $this->_createTree($this->_current_token, '', '');
            $this->_advance();

            return $result;
        }
        // if it's a function call
        if (\preg_match("/^[A-Z0-9\xc0-\xdc\\.]+$/i", $this->_current_token)) {
            $result = $this->_func();

            return $result;
        }

        throw new Excel\Exception\RuntimeException(\sprintf(
            'Syntax error: %s, lookahead: %s, current char: %s',
            $this->_current_token,
            $this->_lookahead,
            $this->_current_char
        ));
    }

    /**
     * It parses a function call. It assumes the following rule:
     * Func -> ( Expr [,Expr]* ).
     *
     *
     * @return mixed The parsed ptg'd tree on success, Excel_PEAR_Error on failure
     */
    private function _func()
    {
        $num_args = 0; // number of arguments received
        $function = \strtoupper($this->_current_token);
        $result   = ''; // initialize result
        $this->_advance();
        $this->_advance();         // eat the "("
        while (')' != $this->_current_token) {
            if ($num_args > 0) {
                if (self::SPREADSHEET_EXCEL_WRITER_COMA == $this->_current_token or
                    self::SPREADSHEET_EXCEL_WRITER_SEMICOLON == $this->_current_token) {
                    $this->_advance();  // eat the "," or ";"
                } else {
                    throw new Excel\Exception\RuntimeException("Syntax error: comma expected in function ${function}, arg #{$num_args}");
                }
                $result2 = $this->_condition();
                $result = $this->_createTree('arg', $result, $result2);
            } else { // first argument
                $result2 = $this->_condition();
                $result = $this->_createTree('arg', '', $result2);
            }
            ++$num_args;
        }
        if (! isset($this->_functions[$function])) {
            throw new Excel\Exception\RuntimeException("Function ${function}() doesn't exist");
        }
        $args = $this->_functions[$function][1];
        // If fixed number of args eg. TIME($i,$j,$k). Check that the number of args is valid.
        if (($args >= 0) and ($args != $num_args)) {
            throw new Excel\Exception\RuntimeException("Incorrect number of arguments in function ${function}() ");
        }

        $result = $this->_createTree($function, $result, $num_args);
        $this->_advance();         // eat the ")"
        return $result;
    }

    /**
     * Creates a tree. In fact an array which may have one or two arrays (sub-trees)
     * as elements.
     *
     *
     * @param mixed $value the value of this node
     * @param mixed $left  the left array (sub-tree) or a final node
     * @param mixed $right the right array (sub-tree) or a final node
     *
     * @return array A tree
     */
    private function _createTree($value, $left, $right)
    {
        return ['value' => $value, 'left' => $left, 'right' => $right];
    }

    /**
     * Builds a string containing the tree in reverse polish notation (What you
     * would use in a HP calculator stack).
     * The following tree:.
     *
     *    +
     *   / \
     *  2   3
     *
     * produces: "23+"
     *
     * The following tree:
     *
     *    +
     *   / \
     *  3   *
     *     / \
     *    6   A1
     *
     * produces: "36A1*+"
     *
     * In fact all operands, functions, references, etc... are written as ptg's
     *
     *
     * @param array $tree the optional tree to convert
     *
     * @return string The tree in reverse polish notation
     */
    public function toReversePolish($tree = [])
    {
        $polish = ''; // the string we are going to return
        if (empty($tree)) { // If it's the first call use _parse_tree
            $tree = $this->_parse_tree;
        }
        if (\is_array($tree['left'])) {
            $converted_tree = $this->toReversePolish($tree['left']);
            $polish .= $converted_tree;
        } elseif ('' != $tree['left']) { // It's a final node
            $converted_tree = $this->_convert($tree['left']);
            $polish .= $converted_tree;
        }
        if (\is_array($tree['right'])) {
            $converted_tree = $this->toReversePolish($tree['right']);
            $polish .= $converted_tree;
        } elseif ('' != $tree['right']) { // It's a final node
            $converted_tree = $this->_convert($tree['right']);
            $polish .= $converted_tree;
        }
        // if it's a function convert it here (so we can set it's arguments)
        if (\preg_match("/^[A-Z0-9\xc0-\xdc\\.]+$/", $tree['value']) and
            ! \preg_match('/^([A-Ia-i]?[A-Za-z])(\d+)$/', $tree['value']) and
            ! \preg_match('/^[A-Ia-i]?[A-Za-z](\\d+)\\.\\.[A-Ia-i]?[A-Za-z](\\d+)$/', $tree['value']) and
            ! \is_numeric($tree['value']) and
            ! isset($this->ptg[$tree['value']])) {
            // left subtree for a function is always an array.
            if ('' != $tree['left']) {
                $left_tree = $this->toReversePolish($tree['left']);
            } else {
                $left_tree = '';
            }
            // add it's left subtree and return.
            return $left_tree . $this->_convertFunction($tree['value'], $tree['right']);
        }
        $converted_tree = $this->_convert($tree['value']);

        $polish .= $converted_tree;

        return $polish;
    }
}
