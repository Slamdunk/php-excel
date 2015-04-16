<?php

use org\bovigo\vfs;

class ExcelTest_MainTest extends PHPUnit_Framework_TestCase
{
    public function testGenerazioneFileBase()
    {
        $vfs = vfs\vfsStream::setup('root', 0770);
        $filename = vfs\vfsStream::url('root/test.xls');

        // $filename = TMP_PATH . '/stock.xls';

        Excel_OLE::$gmmktime = gmmktime('01','01','01','01','01','2000');

        $xls = new Excel_Writer_Workbook($filename);

        $xls->setCustomColor(60, hexdec('7f'), hexdec('7f'), hexdec('7f'));
        $xls->setCustomColor(61, hexdec('e8'), hexdec('e8'), hexdec('e8'));
        $xls->setCustomColor(62, hexdec('cc'), hexdec('cc'), hexdec('cc'));

        $sheet = $xls->addWorksheet('FoglioCustom');
        $sheet->setLandscape();
        $sheet->setMargins(0.2);
        $sheet->hideGridLines();
        $sheet->setColumn(1, 1, 20);

        $header = $xls->addFormat();
        $header->setColor('yellow');
        $header->setBold();
        $header->setSize(15);
        $header->setAlign('center');
        $header->setBgColor(61);

        $sheet->write(1, 1, 'LCR', $header);

        $sheet->freezePanes(array(2, 0));

        $sheet->writeString(2, 1, '0123');

        $format2 = $xls->addFormat();
        $format2->setColor('red');
        $format2->setItalic();
        $format2->setBorder(2);
        $format2->setBorderColor('lime');
        $format2->setNumFormat('#,##0.00');
        $format2->setFgColor(62);
        $format2->setAlign('top');

        $sheet->write(3, 2, '95000', $format2);

        $sheet->write(7, 3, 1.1);
        $sheet->write(8, 3, 2);
        $sheet->writeFormula(9, 3, sprintf('=SUM(%s:%s)', $xls->rowcolToCell(7, 3), $xls->rowcolToCell(8, 3)));

        $xls->close();

        $this->assertSame('43093f883818e44f4dd62f0382d95ecf6b689004', hash('sha1', file_get_contents($filename)));
    }

    /**
     * @dataProvider dataProviderTestIndiceColonnaInNumero
     */
    public function testIndiceColonnaInNumero($indice, $lettera)
    {
        $xls = new Excel_Writer_Workbook();
        $this->assertSame($lettera . '2', $xls->rowcolToCell(1, $indice));

        $this->setExpectedException('Excel_Exception_InvalidArgumentException');

        $xls->rowcolToCell(1, 50000);
    }

    public function dataProviderTestIndiceColonnaInNumero()
    {
        return array(
            array(0, 'A'),
            array(30, 'AE'),
            array(231, 'HX'),
        );
    }
}
