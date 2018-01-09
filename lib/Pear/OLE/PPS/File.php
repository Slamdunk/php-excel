<?php

namespace Slam\Excel\Pear\OLE\PPS;

use Slam\Excel;

/**
 * Class for creating File PPS's for Excel_OLE containers.
 *
 * @author   Xavier Noguer <xnoguer@php.net>
 *
 * @category Structures
 */
class File extends Excel\Pear\OLE\PPS
{
    /**
     * The constructor.
     *
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
            Excel\Pear\OLE::Excel_OLE_PPS_TYPE_FILE,
            null,
            null,
            null,
            null,
            null,
            '',
            []
        );

        $this->_PPS_FILE = Excel\Pear\OLE::getTmpfile();
    }

    /**
     * Append data to PPS.
     *
     *
     * @param string $data The data to append
     */
    public function append($data)
    {
        \fwrite($this->_PPS_FILE, $data);
    }
}
