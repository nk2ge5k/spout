<?php

namespace Box\Spout\Writer\XLSX;

use Box\Spout\Writer\AbstractMultiSheetsWriter;
use Box\Spout\Writer\Common\Options;
use Box\Spout\Writer\Common\Row;
use Box\Spout\Writer\XLSX\Internal\Workbook;

/**
 * Class Writer
 * This class provides base support to write data to XLSX files
 *
 * @package Box\Spout\Writer\XLSX
 */
class Writer extends AbstractMultiSheetsWriter
{
    /** @var string Content-Type value for the header */
    protected static $headerContentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    /** @var Internal\Workbook The workbook for the XLSX file */
    protected $book;

    /**
     * Sets a custom temporary folder for creating intermediate files/folders.
     * This must be set before opening the writer.
     *
     * @api
     * @param string $tempFolder Temporary folder where the files to create the XLSX will be stored
     * @return Writer
     * @throws \Box\Spout\Writer\Exception\WriterAlreadyOpenedException If the writer was already opened
     */
    public function setTempFolder($tempFolder)
    {
        $this->throwIfWriterAlreadyOpened('Writer must be configured before opening it.');

        $this->optionsManager->setOption(Options::TEMP_FOLDER, $tempFolder);
        return $this;
    }

    /**
     * Use inline string to be more memory efficient. If set to false, it will use shared strings.
     * This must be set before opening the writer.
     *
     * @api
     * @param bool $shouldUseInlineStrings Whether inline or shared strings should be used
     * @return Writer
     * @throws \Box\Spout\Writer\Exception\WriterAlreadyOpenedException If the writer was already opened
     */
    public function setShouldUseInlineStrings($shouldUseInlineStrings)
    {
        $this->throwIfWriterAlreadyOpened('Writer must be configured before opening it.');

        $this->optionsManager->setOption(Options::SHOULD_USE_INLINE_STRINGS, $shouldUseInlineStrings);
        return $this;
    }

    /**
     * Configures the write and sets the current sheet pointer to a new sheet.
     *
     * @return void
     * @throws \Box\Spout\Common\Exception\IOException If unable to open the file for writing
     */
    protected function openWriter()
    {
        if (!$this->book) {
            $this->book = new Workbook($this->optionsManager);
            $this->book->addNewSheetAndMakeItCurrent();
        }
    }

    /**
     * @return Internal\Workbook The workbook representing the file to be written
     */
    protected function getWorkbook()
    {
        return $this->book;
    }

    /**
     * @inheritdoc
     */
    protected function addRowToWriter(Row $row)
    {
        $this->throwIfBookIsNotAvailable();
        $this->book->addRowToCurrentWorksheet($row);
    }

    /**
     * Closes the writer, preventing any additional writing.
     *
     * @return void
     */
    protected function closeWriter()
    {
        if ($this->book) {
            $this->book->close($this->filePointer);
        }
    }
}
