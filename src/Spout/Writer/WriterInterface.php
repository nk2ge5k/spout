<?php

namespace Box\Spout\Writer;

use Box\Spout\Common\Exception\IOException;
use Box\Spout\Writer\Common\Entity\Row;

/**
 * Interface WriterInterface
 *
 * @package Box\Spout\Writer
 */
interface WriterInterface
{
    /**
     * Initializes the writer and opens it to accept data.
     * By using this method, the data will be written to a file.
     *
     * @param  string $outputFilePath Path of the output file that will contain the data
     * @return WriterInterface
     * @throws IOException If the writer cannot be opened or if the given path is not writable
     */
    public function openToFile($outputFilePath);

    /**
     * Initializes the writer and opens it to accept data.
     * By using this method, the data will be outputted directly to the browser.
     *
     * @param  string $outputFileName Name of the output file that will contain the data.
     * If a path is passed in, only the file name will be kept
     * @return WriterInterface
     * @throws IOException If the writer cannot be opened
     */
    public function openToBrowser($outputFileName);

    /**
     * Append a row to the end of the stream.
     *
     * @param array|Row $row The row to be appended to the stream
     * @return WriterInterface
     */
    public function addRow($row);

    /**
     * Write given data to the output with a closure function. New data will be appended to the end of the stream.
     *
     * @param \Closure $callback A callback returning a Row object. A new Row object is injected into the callback.
     * @return WriterInterface
     */
    public function withRow(\Closure $callback);

    /**
     * Write a given array of rows to the output. New data will be appended to the end of the stream.
     *
     * @param  Row[] $rows Array of rows be appended to the stream
     * @return WriterInterface
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException If the input param is not valid
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException If the writer has not been opened yet
     * @throws \Box\Spout\Common\Exception\IOException If unable to write data
     */
    public function addRows(array $rows);

    /**
     * Closes the writer. This will close the streamer as well, preventing new data
     * to be written to the file.
     *
     * @return void
     */
    public function close();
}
