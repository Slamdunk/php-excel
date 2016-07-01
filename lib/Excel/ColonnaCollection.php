<?php

final class Excel_ColonnaCollection implements ArrayAccess
{
    private $colonne = array();

    public function __construct(array $colonne)
    {
        foreach ($colonne as $colonna) {
            $this->addColonna($colonna);
        }
    }

    private function addColonna(Excel_ColonnaInterface $colonna)
    {
        $this->colonne[$colonna->getChiave()] = $colonna;
    }

    public function offsetSet($offset, $value)
    {
        throw new Excel_Exception_RuntimeException('Collezione non modificabile');
    }

    public function offsetExists($offset)
    {
        return isset($this->colonne[$offset]);
    }

    public function offsetUnset($offset)
    {
        throw new Excel_Exception_RuntimeException('Collezione non modificabile');
    }

    public function offsetGet($offset)
    {
        return isset($this->colonne[$offset]) ? $this->colonne[$offset] : null;
    }
}
