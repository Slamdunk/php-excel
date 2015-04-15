<?php

use org\bovigo\vfs;

class ExcelTest_MainTest extends PHPUnit_Framework_TestCase
{
    public function testGenerazioneFileBase()
    {
        $vfs = vfs\vfsStream::setup('root', 0770);
        $filename = vfs\vfsStream::url('root/test.xls');
        
        $filename = __DIR__ . '/stock.xls';
        
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
        $header->setBold();
        $header->setSize(15);
        $header->setAlign('center');

        $sheet->write(1, 1, 'LCR', $header);

        $sheet->freezePanes(array(2, 0));

        $sheet->writeString(2, 1, '0123');
        
        // $format2 = $xls->addFormat();
        
        $xls->close();

        $this->assertSame('d2551ceac5226138770322d59cdb3719337c54fb', hash('sha1', file_get_contents($filename)));
    }
}
