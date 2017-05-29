<?php

namespace Box\Spout\Writer\Common\Internal;

use Box\Spout\Writer\Common\Row;

/**
 * Interface WorkbookInterface
 *
 * @package Box\Spout\Writer\Common\Internal
 */
interface WorkbookInterface
{
    /**
     * Creates a new sheet in the workbook. The current sheet remains unchanged.
     *
     * @return WorksheetInterface The created sheet
     * @throws \Box\Spout\Common\Exception\IOException If unable to open the sheet for writing
     */
    public function addNewSheet();

    /**
     * Creates a new sheet in the workbook and make it the current sheet.
     * The writing will resume where it stopped (i.e. data won't be truncated).
     *
     * @return WorksheetInterface The created sheet
     * @throws \Box\Spout\Common\Exception\IOException If unable to open the sheet for writing
     */
    public function addNewSheetAndMakeItCurrent();

    /**
     * @return WorksheetInterface[] All the workbook's sheets
     */
    public function getWorksheets();

    /**
     * Returns the current sheet
     *
     * @return WorksheetInterface The current sheet
     */
    public function getCurrentWorksheet();

    /**
     * Sets the given sheet as the current one. New data will be written to this sheet.
     * The writing will resume where it stopped (i.e. data won't be truncated).
     *
     * @param \Box\Spout\Writer\Common\Sheet $sheet The "external" sheet to set as current
     * @return void
     * @throws \Box\Spout\Writer\Exception\SheetNotFoundException If the given sheet does not exist in the workbook
     */
    public function setCurrentSheet($sheet);

    /**
     * Adds data to the current sheet.
     * If shouldCreateNewSheetsAutomatically option is set to true, it will handle pagination
     * with the creation of new worksheets if one worksheet has reached its maximum capicity.
     *
     * @param Row $row The row to be appended to the sheet
     * @return void
     */
    public function addRowToCurrentWorksheet(Row $row);

    /**
     * Closes the workbook and all its associated sheets.
     * All the necessary files are written to disk and zipped together to create the ODS file.
     * All the temporary files are then deleted.
     *
     * @param resource $finalFilePointer Pointer to the ODS that will be created
     * @return void
     */
    public function close($finalFilePointer);
}
