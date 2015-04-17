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

        $this->_PPS_FILE = Excel_OLE::getTmpfile();
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
        fwrite($this->_PPS_FILE, $data);
    }
}
