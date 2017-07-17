<?php

namespace FiiSoft\Test\Tools\PHPExcel\Worksheet;

use FiiSoft\Tools\PHPExcel\Worksheet\SortedPHPExcelWorksheetRowIterator;
use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Worksheet;
use PHPExcel_Worksheet_Row;
use PHPExcel_Worksheet_RowCellIterator;

class SortedPHPExcelWorksheetRowIteratorTest extends \PHPUnit_Framework_TestCase
{
    public function test_iterate_sheet_by_particular_column()
    {
        $filePath = implode(DIRECTORY_SEPARATOR, [__DIR__, '..', '..', '..', 'files', 'sheet.xlsx']);
        self::assertFileExists($filePath);
        
        $excel = PHPExcel_IOFactory::load($filePath);
        self::assertInstanceOf(PHPExcel::class, $excel);
        
        $sheet = $excel->getSheetByName('Sheet');
        self::assertInstanceOf(PHPExcel_Worksheet::class, $sheet);
        
        $it = new SortedPHPExcelWorksheetRowIterator($sheet, 'B', 4, 21);
        self::assertSame('B', $it->sortedByColumn());
    
        $index = 0;
        $keysNumbers = [4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21];
        $rowsNumbers = [12,9,6,10,17,14,19,4,15,18,11,7,5,20,16,21,13,8];
        $valuesInColA = [9,6,3,7,14,11,16,1,12,15,8,4,2,17,13,18,10,5];
        $valuesInColB = [1,2,3,4,4,5,6,7,7,8,9,10,11,11,12,13,14,15];
        
        while ($it->valid()) {
            self::assertSame($keysNumbers[$index], $it->key());
            self::assertSame($rowsNumbers[$index], $it->row());
            
            $row = $it->current();
            self::assertInstanceOf(PHPExcel_Worksheet_Row::class, $row);
    
            /* @var $cells PHPExcel_Worksheet_RowCellIterator */
            $cells = $row->getCellIterator();
            self::assertInstanceOf(PHPExcel_Worksheet_RowCellIterator::class, $cells);
            
            self::assertSame($valuesInColA[$index], (int) $cells->current()->getValue());
            
            $cells->next();
            self::assertSame($valuesInColB[$index], (int) $cells->current()->getValue());
            
            $it->next();
            ++$index;
        }
    }
}
