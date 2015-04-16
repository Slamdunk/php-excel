<?php

/**
 * Class for creating File PPS's for Excel_OLE containers
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 *
 * @category Structures
 */
class Excel_OLE_PPS_File extends Excel_OLE_PPS
{
    /**
     * The temporary dir for storing the Excel_OLE file
     *
     * @var string
     */
    private $_tmp_dir;

    /**
     * The constructor
     *
     * @access public
     *
     * @param string $name The name of the file (in Unicode)
     *
     * @see Excel_OLE::Asc2Ucs()
     */
    public function __construct($name)
    {
        $this->_tmp_dir = sys_get_temp_dir();

        parent::__construct(
            null,
            $name,
            Excel_OLE::Excel_OLE_PPS_TYPE_FILE,
            null,
            null,
            null,
            null,
            null,
            '',
            array()
        );
    }

    /**
     * Sets the temp dir used for storing the Excel_OLE file
     *
     * @access public
     *
     * @param string $dir The dir to be used as temp dir
     *
     * @return true if given dir is valid, false otherwise
     */
    public function setTempDir($dir)
    {
        if (is_dir($dir)) {
            $this->_tmp_dir = $dir;

            return true;
        }

        return false;
    }

    /**
     * Initialization method. Has to be called right after Excel_OLE_PPS_File().
     *
     * @access public
     *
     * @return mixed true on success. PEAR_Error on failure
     */
    public function init()
    {
        $this->_tmp_filename = tempnam($this->_tmp_dir, "Excel_OLE_PPS_File");
        $fh = @fopen($this->_tmp_filename, "w+b");
        if ($fh == false) {
            throw new Excel_Exception_RuntimeException("Can't create temporary file: " . $this->_tmp_filename);
        }
        $this->_PPS_FILE = $fh;
        if ($this->_PPS_FILE) {
            fseek($this->_PPS_FILE, 0);
        }

        return true;
    }

    /**
     * Append data to PPS
     *
     * @access public
     *
     * @param string $data The data to append
     */
    public function append($data)
    {
        if ($this->_PPS_FILE) {
            fwrite($this->_PPS_FILE, $data);
        } else {
            $this->_data .= $data;
        }
    }

    /**
     * Returns a stream for reading this file using fread() etc.
     *
     * @return resource a read-only stream
     */
    public function getStream()
    {
        $this->ole->getStream($this);
    }
}
