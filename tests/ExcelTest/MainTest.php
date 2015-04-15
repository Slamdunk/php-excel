<?php

use org\bovigo\vfs;

class ExcelTest_MainTest extends PHPUnit_Framework_TestCase
{
    public function testGenerazioneFileBase()
    {
        $vfs = vfs\vfsStream::setup('root', 0770);
        $filename = vfs\vfsStream::url('root/test.xls');
        
        $filename = TMP_PATH . '/stock.xls';
        
        Excel_OLE::$gmmktime = gmmktime('01','01','01','01','01','2000');

        $xls = new Excel_Writer($filename);

        $xls->setCustomColor( 60, hexdec( '7f' ), hexdec( '7f' ), hexdec( '7f' ) );
        $xls->setCustomColor( 61, hexdec( 'e8' ), hexdec( 'e8' ), hexdec( 'e8' ) );
        $xls->setCustomColor( 62, hexdec( 'cc' ), hexdec( 'cc' ), hexdec( 'cc' ) );

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
        
        $xls->close();

        $this->assertSame('d2551ceac5226138770322d59cdb3719337c54fb', hash('sha1', file_get_contents($filename)));
    }
}
