<?php

namespace Arodiss\XlsBundle\Xls\Reader;

interface ReaderInterface
{
    /**
     * @param string $path
     * @return array
     */
    public function readAll($path);

    /**
     * @param $path
     * @param int $startRow
     * @param int $size
     * @param bool $withEmptyRows
     * @return array
     */
    public function getRowsChunk($path, $startRow = 1, $size = 65000, $withEmptyRows = false);

    /**
     * @param string $path
     * @return \Iterator
     */
    public function getReadIterator($path);

    /**
     * @param string $path
     * @param int $maxCountEmptyRows
     * @param bool $withEmptyRows
     * @return int
     */
    public function getRowsNumber($path, $maxCountEmptyRows = null, $withEmptyRows = false);

    /**
     * @param string $path
     * @return array
     */
    public function getHeaderRow($path);
}
