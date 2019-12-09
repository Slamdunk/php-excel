<?php

declare(strict_types=1);

namespace Slam\Excel\Tests\Helper;

use org\bovigo\vfs;
use PHPUnit\Framework\TestCase;
use Slam\Excel;

final class MainTest extends TestCase
{
    private $vfs;
    private $filename;
    private $xls;

    protected function setUp(): void
    {
        Excel\Pear\OLE::$gmmktime = \gmmktime(1, 1, 1, 1, 1, 2000);

        $this->vfs      = vfs\vfsStream::setup('root', 0770);
        $this->filename = vfs\vfsStream::url('root/test.xls');

        // $this->filename = __DIR__ . '/stock.xls';

        $this->xls = new Excel\Pear\Writer\Workbook($this->filename);
    }

    public function testBaseCreation(): void
    {
        $this->xls->setCustomColor(60, \hexdec('7f'), \hexdec('7f'), \hexdec('7f'));
        $this->xls->setCustomColor(61, \hexdec('e8'), \hexdec('e8'), \hexdec('e8'));
        $this->xls->setCustomColor(62, \hexdec('cc'), \hexdec('cc'), \hexdec('cc'));

        $sheet = $this->xls->addWorksheet('CustomSheet');
        $sheet->setLandscape();
        $sheet->setMargins(0.2);
        $sheet->hideGridLines();
        $sheet->setColumn(1, 1, 20);

        $header = $this->xls->addFormat();
        $header->setColor('yellow');
        $header->setBold();
        $header->setSize(15);
        $header->setAlign('center');
        $header->setBgColor(61);

        $sheet->write(1, 1, 'LCR', $header);

        $sheet->freezePanes([2, 0]);

        $sheet->writeString(2, 1, '0123');

        $format2 = $this->xls->addFormat();
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
        $sheet->writeFormula(9, 3, \sprintf('=SUM(%s:%s)', $this->xls->rowcolToCell(7, 3), $this->xls->rowcolToCell(8, 3)));

        $this->xls->close();

        self::assertLessThan(Excel\Pear\OLE::Excel_OLE_DATA_SIZE_SMALL, \filesize($this->filename));
        self::assertSame('9ab414a26fdb4bfb5a6d65c9c214ddc70b5a5464', \hash_file('sha1', $this->filename));
    }

    public function testBigFiles(): void
    {
        $sheet = $this->xls->addWorksheet('CustomSheet');

        for ($i = 0; $i < 1000; ++$i) {
            $sheet->writeString($i, 1, 'foobar' . $i);
        }

        $this->xls->close();

        self::assertGreaterThan(Excel\Pear\OLE::Excel_OLE_DATA_SIZE_SMALL, \filesize($this->filename));
        self::assertSame('f95f7ff2fb1ffb98cef820ecd73312c9a43b9662', \hash_file('sha1', $this->filename));
    }

    /**
     * @dataProvider dataProviderTestColumnIndexInInumber
     */
    public function testColumnIndexInInumber(int $index, string $letter): void
    {
        self::assertSame($letter . '2', $this->xls->rowcolToCell(1, $index));

        $this->expectException(Excel\Exception\InvalidArgumentException::class);

        $this->xls->rowcolToCell(1, 50000);
    }

    public function dataProviderTestColumnIndexInInumber()
    {
        return [
            [0, 'A'],
            [30, 'AE'],
            [231, 'HX'],
        ];
    }
}
