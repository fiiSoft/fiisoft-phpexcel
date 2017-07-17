<?php

namespace FiiSoft\Tools\PHPExcel\Worksheet;

use BadMethodCallException;
use InvalidArgumentException;
use LogicException;
use PHPExcel_Cell;
use PHPExcel_Exception;
use PHPExcel_Worksheet;
use PHPExcel_Worksheet_ColumnCellIterator;
use PHPExcel_Worksheet_Row;
use PHPExcel_Worksheet_RowIterator;

final class SortedPHPExcelWorksheetRowIterator extends PHPExcel_Worksheet_RowIterator
{
    /** @var PHPExcel_Worksheet */
    private $sheet;
    
    /** @var string */
    private $sortByColumn;
    
    /** @var int */
    private $firstRow;
    
    /** @var int */
    private $lastRow = 0;
    
    /** @var int */
    private $startIndex = 0;
    
    /** @var int */
    private $stopIndex = 0;
    
    /** @var int */
    private $index = 0;
    
    /** @var array */
    private $indexes = [];
    
    /** @noinspection PhpMissingParentConstructorInspection */
    /** @noinspection MagicMethodsValidityInspection */
    
    /**
     * @param PHPExcel_Worksheet $sheet
     * @param string $sortByColumn
     * @param integer $startFromRow
     * @param integer $endOnRow
     * @throws InvalidArgumentException
     * @throws LogicException
     */
    public function __construct(PHPExcel_Worksheet $sheet, $sortByColumn, $startFromRow = 1, $endOnRow = null)
    {
        if (!is_string($sortByColumn) || $sortByColumn === '') {
            throw new InvalidArgumentException('Invalid parameter sortByColumn - a non-empty string is required');
        }
    
        if (!is_int($startFromRow) || $startFromRow < 1) {
            throw new InvalidArgumentException('Invalid parameter startFromRow - it must be an integer not less then 1');
        }
        
        $highestRow = $sheet->getHighestRow($sortByColumn);
        
        if ($endOnRow === null) {
            $endOnRow = $highestRow;
        } elseif (!is_int($endOnRow) || $endOnRow < $startFromRow || $endOnRow > $highestRow) {
            throw new InvalidArgumentException(
                'Invalid parameter endOnRow - it must be an integer between '.$startFromRow.' and '.$highestRow
            );
        }
        
        if ($startFromRow > $endOnRow) {
            throw new InvalidArgumentException(
                'Invalid parameter startFromRow - it must be an integer between 1 and '.$endOnRow.' (last row in range).'
            );
        }
        
        $this->sheet = $sheet;
        $this->sortByColumn = $sortByColumn;
        $this->firstRow = $startFromRow;
        
        $values = [];
        $i = $startFromRow;
        
        $iterator = new PHPExcel_Worksheet_ColumnCellIterator($sheet, $sortByColumn, $startFromRow, $endOnRow);
        /* @var $cell PHPExcel_Cell */
        foreach ($iterator as $cell) {
            $values[] = $cell->getValue();
            $this->indexes[] = $i++;
        }
        
        array_multisort($values, $this->indexes);
        
        $this->stopIndex = count($this->indexes) - 1;
        $this->lastRow = $this->firstRow + $this->stopIndex;
    }
    
    /** @noinspection MagicMethodsValidityInspection */
    public function __destruct()
    {
        unset($this->sheet);
        $this->indexes = [];
    }
    
    /**
     * @param integer $startRow
     * @throws InvalidArgumentException
     * @return $this fluent interface
     */
    public function resetStart($startRow = 1)
    {
        if (is_int($startRow)) {
            if ($startRow < $this->firstRow) {
                $startRow = $this->firstRow;
            } elseif ($startRow > $this->lastRow) {
                throw new InvalidArgumentException(
                    'Invalid param startRow, it must be an integer in the range <'
                    .$this->firstRow.','.$this->lastRow.'>'
                );
            }
        } else {
            throw new InvalidArgumentException('Invalid param startRow, it must be an integer');
        }
        
        $this->index = $this->startIndex = $startRow - $this->firstRow;
        
        if ($this->index > $this->stopIndex) {
            $this->stopIndex = $this->lastRow - $this->firstRow;
        }
        
        return $this;
    }
    
    /**
     * @param integer $endRow
     * @throws InvalidArgumentException
     * @return $this fluent interface
     */
    public function resetEnd($endRow = null)
    {
        if ($endRow === null) {
            $endRow = $this->lastRow;
        } elseif (is_int($endRow)) {
            if ($endRow < $this->firstRow || $endRow > $this->lastRow) {
                throw new InvalidArgumentException(
                    'Invalid param endRow, it must be an integer in the range <'
                    .$this->firstRow.','.$this->lastRow.'>'
                );
            }
        } else {
            throw new InvalidArgumentException('Invalid param endRow, it must be an integer');
        }
        
        $this->stopIndex = $endRow - $this->firstRow;
        
        if ($this->stopIndex < $this->startIndex) {
            $this->startIndex = 0;
        }
        
        return $this;
    }
    
    /**
     * @param int $row
     * @throws PHPExcel_Exception
     * @throws InvalidArgumentException
     * @return $this fluent interface
     */
    public function seek($row = 1)
    {
        if (is_int($row)) {
            $min = $this->startIndex + $this->firstRow;
            $max = $this->stopIndex + $this->firstRow;
            if ($row < $min || $row > $max) {
                throw new InvalidArgumentException(
                    'Invalid parameter row, it must be an integer in the range <'.$min.','.$max.'>'
                );
            }
        } else {
            throw new InvalidArgumentException('Invalid param row, it must be an integer');
        }
        
        $this->index = $row - $this->firstRow;
        
        return $this;
    }
    
    /**
     * @return void
     */
    public function rewind()
    {
        $this->index = $this->startIndex;
    }
    
    /**
     * @throws BadMethodCallException
     * @return PHPExcel_Worksheet_Row
     */
    public function current()
    {
        if (!$this->valid()) {
            throw new BadMethodCallException('Iterator is over, restart it or use seek to fetch current value');
        }
        
        return new PHPExcel_Worksheet_Row($this->sheet, $this->row());
    }
    
    /**
     * @return int number of row in sorted sheet
     */
    public function key()
    {
        return $this->index + $this->firstRow;
    }
    
    /**
     * @return int original number of row in decorated (unsorted) sheet
     */
    public function row()
    {
        return $this->indexes[$this->index];
    }
    
    /**
     * @throws PHPExcel_Exception
     * @return void
     */
    public function next()
    {
        if ($this->index > $this->stopIndex) {
            throw new PHPExcel_Exception('Iterator has reached its end, restart it before iterate again');
        }
        
        ++$this->index;
    }
    
    /**
     * @throws PHPExcel_Exception
     * @return void
     */
    public function prev()
    {
        if ($this->index === $this->startIndex) {
            throw new PHPExcel_Exception('Row is already at the beginning of range');
        }
        
        --$this->index;
    }
    
    /**
     * @return bool
     */
    public function valid()
    {
        return $this->index <= $this->stopIndex && $this->index >= $this->startIndex;
    }
    
    /**
     * @return string
     */
    public function sortedByColumn()
    {
        return $this->sortByColumn;
    }
}