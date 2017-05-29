<?php

namespace Box\Spout\Writer\Common\Internal;

use Box\Spout\Writer\Common\Row;

/**
 * Interface WorksheetInterface
 *
 * @package Box\Spout\Writer\Common\Internal
 */
interface WorksheetInterface
{
    /**
     * @return \Box\Spout\Writer\Common\Sheet The "external" sheet
     */
    public function getExternalSheet();

    /**
     * @return int The index of the last written row
     */
    public function getLastWrittenRowIndex();

    /**
     * Adds data to the worksheet.
     *
     * @param Row $row The row to be added
     * @return void
     */
    public function addRow(Row $row);

    /**
     * Closes the worksheet
     *
     * @return void
     */
    public function close();
}
