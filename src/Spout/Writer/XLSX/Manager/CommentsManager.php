<?php

namespace Box\Spout\Writer\XLSX\Manager;

use Box\Spout\Common\Exception\IOException;
use Box\Spout\Common\Helper\Escaper;

/**
 * Class commentsManager
 * This class provides functions to write cell comments
 */
class CommentsManager
{
    const COMMENTS_FILE_NAME = 'comments1.xml';

    const COMMENTS_XML_FILE_FIRST_PART_HEADER = <<<'EOD'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Author</author></authors>
<commentList>
EOD;

    /** @var resource Pointer to the comments.xml file */
    protected $commentsFilePointer;

    /** @var int Number of shared strings already written */
    protected $numComments = 0;

    /** @var Escaper\XLSX Strings escaper */
    protected $stringsEscaper;

    /**
     * @param string $xlFolder Path to the "xl" folder
     * @param Escaper\XLSX $stringsEscaper Strings escaper
     */
    public function __construct($xlFolder, $stringsEscaper, $sheetIndex)
    {
        $commentsFilePath = $xlFolder . '/comments' . $sheetIndex . ".xml";
        $this->commentsFilePointer = \fopen($commentsFilePath, 'w');

        $this->throwIfCommentsFilePointerIsNotAvailable();

        $header = self::COMMENTS_XML_FILE_FIRST_PART_HEADER;
        \fwrite($this->commentsFilePointer, $header);

        $this->stringsEscaper = $stringsEscaper;
    }

    /**
     * Checks if the book has been created. Throws an exception if not created yet.
     *
     * @throws \Box\Spout\Common\Exception\IOException If the sheet data file cannot be opened for writing
     * @return void
     */
    protected function throwIfCommentsFilePointerIsNotAvailable()
    {
        if (!$this->commentsFilePointer) {
            throw new IOException('Unable to open comments file for writing.');
        }
    }

    /**
     * Writes the given comment into the comments.xml file.
     * Starting and ending whitespaces are preserved.
     *
     * @param Comment $comment
     * @param string $cellReference The cell reference (eg 'B3')
     */
    public function writeComment(Comment $comment, $cellReference)
    {
        $string = $comment->getValue();
        \fwrite($this->commentsFilePointer, '<comment><t xml:space="preserve">' . $this->stringsEscaper->escape($string) . '</t></comment>');
        $this->numcomments++;
    }

    /**
     * Finishes writing the data in the comments.xml file and closes the file.
     *
     * @return void
     */
    public function close()
    {
        if (!\is_resource($this->commentsFilePointer)) {
            return;
        }

        \fwrite($this->commentsFilePointer, '</commentList></comments>');
        \fclose($this->commentsFilePointer);
    }
}
