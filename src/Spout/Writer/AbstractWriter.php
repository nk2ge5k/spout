<?php

namespace Box\Spout\Writer;

use Box\Spout\Common\Exception\InvalidArgumentException;
use Box\Spout\Common\Exception\IOException;
use Box\Spout\Common\Exception\SpoutException;
use Box\Spout\Common\Helper\FileSystemHelper;
use Box\Spout\Writer\Common\Cell;
use Box\Spout\Writer\Common\Manager\OptionsManagerInterface;
use Box\Spout\Writer\Common\Options;
use Box\Spout\Writer\Common\Row;
use Box\Spout\Writer\Exception\WriterAlreadyOpenedException;
use Box\Spout\Writer\Exception\WriterNotOpenedException;
use Box\Spout\Writer\Style\Style;

/**
 * Class AbstractWriter
 *
 * @package Box\Spout\Writer
 * @abstract
 */
abstract class AbstractWriter implements WriterInterface
{
    /** @var string Path to the output file */
    protected $outputFilePath;

    /** @var resource Pointer to the file/stream we will write to */
    protected $filePointer;

    /** @var bool Indicates whether the writer has been opened or not */
    protected $isWriterOpened = false;

    /** @var \Box\Spout\Common\Helper\GlobalFunctionsHelper Helper to work with global functions */
    protected $globalFunctionsHelper;

    /** @var \Box\Spout\Writer\Common\Manager\OptionsManagerInterface Writer options manager */
    protected $optionsManager;

    /** @var string Content-Type value for the header - to be defined by child class */
    protected static $headerContentType;

    /**
     * Opens the streamer and makes it ready to accept data.
     *
     * @return void
     * @throws \Box\Spout\Common\Exception\IOException If the writer cannot be opened
     */
    abstract protected function openWriter();

    /**
     * Adds data to the currently opened writer.
     *
     * @param Row $row The row to appended to the stream
     * @return void
     */
    abstract protected function addRowToWriter(Row $row);

    /**
     * Closes the streamer, preventing any additional writing.
     *
     * @return void
     */
    abstract protected function closeWriter();

    /**
     * @param \Box\Spout\Writer\Common\Manager\OptionsManagerInterface $optionsManager
     */
    public function __construct(OptionsManagerInterface $optionsManager)
    {
        $this->optionsManager = $optionsManager;
    }

    /**
     * Sets the default styles for all rows added with "addRow".
     * Overriding the default style instead of using "addRowWithStyle" improves performance by 20%.
     * @see https://github.com/box/spout/issues/272
     *
     * @param Style $defaultStyle
     * @return AbstractWriter
     */
    public function setDefaultRowStyle($defaultStyle)
    {
        $this->optionsManager->setOption(Options::DEFAULT_ROW_STYLE, $defaultStyle);
        return $this;
    }

    /**
     * @param \Box\Spout\Common\Helper\GlobalFunctionsHelper $globalFunctionsHelper
     * @return AbstractWriter
     */
    public function setGlobalFunctionsHelper($globalFunctionsHelper)
    {
        $this->globalFunctionsHelper = $globalFunctionsHelper;
        return $this;
    }

    /**
     * Inits the writer and opens it to accept data.
     * By using this method, the data will be written to a file.
     *
     * @api
     * @param  string $outputFilePath Path of the output file that will contain the data
     * @return AbstractWriter
     * @throws \Box\Spout\Common\Exception\IOException If the writer cannot be opened or if the given path is not writable
     */
    public function openToFile($outputFilePath)
    {
        $this->outputFilePath = $outputFilePath;

        $this->filePointer = $this->globalFunctionsHelper->fopen($this->outputFilePath, 'wb+');
        $this->throwIfFilePointerIsNotAvailable();

        $this->openWriter();
        $this->isWriterOpened = true;

        return $this;
    }

    /**
     * Inits the writer and opens it to accept data.
     * By using this method, the data will be outputted directly to the browser.
     *
     * @codeCoverageIgnore
     *
     * @api
     * @param  string $outputFileName Name of the output file that will contain the data. If a path is passed in, only the file name will be kept
     * @return AbstractWriter
     * @throws \Box\Spout\Common\Exception\IOException If the writer cannot be opened
     */
    public function openToBrowser($outputFileName)
    {
        $this->outputFilePath = $this->globalFunctionsHelper->basename($outputFileName);

        $this->filePointer = $this->globalFunctionsHelper->fopen('php://output', 'w');
        $this->throwIfFilePointerIsNotAvailable();

        // Clear any previous output (otherwise the generated file will be corrupted)
        // @see https://github.com/box/spout/issues/241
        $this->globalFunctionsHelper->ob_end_clean();

        // Set headers
        $this->globalFunctionsHelper->header('Content-Type: ' . static::$headerContentType);
        $this->globalFunctionsHelper->header('Content-Disposition: attachment; filename="' . $this->outputFilePath . '"');

        /*
         * When forcing the download of a file over SSL,IE8 and lower browsers fail
         * if the Cache-Control and Pragma headers are not set.
         *
         * @see http://support.microsoft.com/KB/323308
         * @see https://github.com/liuggio/ExcelBundle/issues/45
         */
        $this->globalFunctionsHelper->header('Cache-Control: max-age=0');
        $this->globalFunctionsHelper->header('Pragma: public');

        $this->openWriter();
        $this->isWriterOpened = true;

        return $this;
    }

    /**
     * Checks if the pointer to the file/stream to write to is available.
     * Will throw an exception if not available.
     *
     * @return void
     * @throws \Box\Spout\Common\Exception\IOException If the pointer is not available
     */
    protected function throwIfFilePointerIsNotAvailable()
    {
        if (!$this->filePointer) {
            throw new IOException('File pointer has not be opened');
        }
    }

    /**
     * Checks if the writer has already been opened, since some actions must be done before it gets opened.
     * Throws an exception if already opened.
     *
     * @param string $message Error message
     * @return void
     * @throws \Box\Spout\Writer\Exception\WriterAlreadyOpenedException If the writer was already opened and must not be.
     */
    protected function throwIfWriterAlreadyOpened($message)
    {
        if ($this->isWriterOpened) {
            throw new WriterAlreadyOpenedException($message);
        }
    }

    /**
     * @inheritdoc
     */
    public function addRow($row)
    {
        if (!is_array($row) && !$row instanceof Row) {
            throw new InvalidArgumentException('addRow accepts an array with scalar values or a Row object');
        }

        if (is_array($row)) {
            $row = $this->createRowFromArray($row, null);
        }

        $this->applyDefaultRowStyle($row);

        if ($this->isWriterOpened) {
            try {
                $this->addRowToWriter($row);
            } catch (SpoutException $e) {
                // if an exception occurs while writing data,
                // close the writer and remove all files created so far.
                $this->closeAndAttemptToCleanupAllFiles();

                // re-throw the exception to alert developers of the error
                throw $e;
            }
        } else {
            throw new WriterNotOpenedException('The writer needs to be opened before adding row.');
        }

        return $this;
    }

    /**
     * @inheritdoc
     */
    public function addRowWithStyle($row, $style)
    {
        if (!is_array($row) && !$row instanceof Row) {
            throw new InvalidArgumentException('addRowWithStyle accepts an array with scalar values or a Row object');
        }

        if (!$style instanceof Style) {
            throw new InvalidArgumentException('The "$style" argument must be a Style instance and cannot be NULL.');
        }

        if(is_array($row)) {
            $row = $this->createRowFromArray($row, $style);
        }

        $this->addRow($row);

        return $this;
    }

    /**
     * @param array $dataRows
     * @param Style|null $style
     * @return Row
     */
    protected function createRowFromArray(array $dataRows, Style $style = null)
    {
        $row = (new Row())->setCells(array_map(function ($value) {
            if ($value instanceof Cell) {
                return $value;
            }
            return new Cell($value);
        }, $dataRows));

        if($style !== null) {
            $row->setStyle($style);
        }

        return $row;
    }

    /**
     * @inheritdoc
     */
    public function addRows(array $dataRows)
    {
        if (!empty($dataRows)) {
            $firstRow = reset($dataRows);
            if (!is_array($firstRow) && !$firstRow instanceof Row) {
                throw new InvalidArgumentException('The input should be an array of arrays or row objects');
            }

            foreach ($dataRows as $dataRow) {
                $this->addRow($dataRow);
            }
        }

        return $this;
    }

    /**
     * @inheritdoc
     */
    public function addRowsWithStyle(array $dataRows, $style)
    {
        if (!$style instanceof Style) {
            throw new InvalidArgumentException('The "$style" argument must be a Style instance and cannot be NULL.');
        }

        $this->addRows(array_map(function($row) use ($style) {
            if(is_array($row)) {
                return $this->createRowFromArray($row, $style);
            } elseif ($row instanceof Row) {
                return $row;
            } else {
                throw new InvalidArgumentException();
            }
        }, $dataRows));

        return $this;
    }

    /**
     * @param Row $row
     * @return $this
     */
    private function applyDefaultRowStyle(Row $row)
    {
        $defaultRowStyle = $this->optionsManager->getOption(Options::DEFAULT_ROW_STYLE);
        $row->applyStyle($defaultRowStyle);
        return $this;
    }

    /**
     * Closes the writer. This will close the streamer as well, preventing new data
     * to be written to the file.
     *
     * @api
     * @return void
     */
    public function close()
    {
        if (!$this->isWriterOpened) {
            return;
        }

        $this->closeWriter();

        if (is_resource($this->filePointer)) {
            $this->globalFunctionsHelper->fclose($this->filePointer);
        }

        $this->isWriterOpened = false;
    }

    /**
     * Closes the writer and attempts to cleanup all files that were
     * created during the writing process (temp files & final file).
     *
     * @return void
     */
    private function closeAndAttemptToCleanupAllFiles()
    {
        // close the writer, which should remove all temp files
        $this->close();

        // remove output file if it was created
        if ($this->globalFunctionsHelper->file_exists($this->outputFilePath)) {
            $outputFolderPath = dirname($this->outputFilePath);
            $fileSystemHelper = new FileSystemHelper($outputFolderPath);
            $fileSystemHelper->deleteFile($this->outputFilePath);
        }
    }
}
