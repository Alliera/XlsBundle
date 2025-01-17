<?php

namespace Arodiss\XlsBundle\Xls\Reader;

use Arodiss\XlsBundle\Iterator\NestingDiscloseIterator;
use Arodiss\XlsBundle\Iterator\StringifyIterator;
use Arodiss\XlsBundle\Filter\ReadFilter;
use Arodiss\XlsBundle\Filter\RowFilter;

class PhpofficeReader extends ReaderAbstract implements ReaderInterface
{
    /** \PHPExcel_Reader_Abstract[] */
    private $worksheets = [];

    /** {@inheirtdoc} */
    public function readAll($path)
    {
        return $this->getExcel($path)->toArray();
    }

    /** {@inheritdoc} */
    public function getRowsChunk($path, $startRow = 1, $size = 65000, $withEmptyRows = false)
    {
        return RowFilter::clearEmptyRows($this->getWorksheetPart($path, $startRow, $size)->toArray());
    }

    /** {@inheritdoc} */
    public function getReadIterator($path)
    {
        return new StringifyIterator(new NestingDiscloseIterator($this->getExcel($path)->getRowIterator()));
    }

    /** {@inheritdoc} */
    public function getRowsNumber($path, $maxCountEmptyRows = null, $withEmptyRows = false)
    {
        return $this->getExcel($path)->getHighestRow();
    }

    /**
     * @param string $path
     * @return \PHPExcel_Worksheet
     * @throws \PHPExcel_Reader_Exception
     */
    protected function getExcel($path)
    {
        /** @var \PHPExcel $excel */
        return $this->getWorksheet($path);
    }

    /**
     * @param $path
     * @param $startRow
     * @param $size
     * @return \PHPExcel_Worksheet
     * @throws \PHPExcel_Reader_Exception
     */
    protected function getWorksheetPart($path, $startRow, $size)
    {
        return $this->getWorksheet($path, new ReadFilter($startRow, $size));
    }

    /**
     * @param $path
     * @param \PHPExcel_Reader_IReadFilter|null $readFilter
     * @return mixed
     */
    protected function getWorksheet($path, \PHPExcel_Reader_IReadFilter $readFilter = null)
    {
        if (false === isset($this->worksheets[$path]) || $readFilter) {
            $reader = $this->createReaderForFile($path);
            $reader->setReadDataOnly(true);
            if ($readFilter) {
                $reader->setReadFilter($readFilter);
            }
            $this->worksheets[$path] = $reader->load($path)->getActiveSheet();
        }

        return $this->worksheets[$path];
    }

    /**
     * @param string $path
     * @return \PHPExcel_Reader_Abstract
     */
    protected function createReaderForFile($path)
    {
        return \PHPExcel_IOFactory::createReaderForFile($path);
    }
}
