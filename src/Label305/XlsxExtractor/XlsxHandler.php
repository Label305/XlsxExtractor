<?php

namespace Label305\XlsxExtractor;

use DirectoryIterator;
use DOMDocument;

abstract class XlsxHandler extends ZipHandler {

    /**
     * Defaults to sys_get_temp_dir()
     *
     * @var string the tmp dir location
     */
    protected $temporaryDirectory;

    /**
     * Sets the temporary directory to the system
     */
    function __construct()
    {
        $this->setTemporaryDirectory(sys_get_temp_dir());
    }

    /**
     * @return string|null
     */
    public function getTemporaryDirectory(): ?string
    {
        return $this->temporaryDirectory;
    }

    /**
     * @param string|null $temporaryDirectory
     * @return $this
     */
    public function setTemporaryDirectory(?string $temporaryDirectory)
    {
        $this->temporaryDirectory = $temporaryDirectory;
        return $this;
    }

    /**
     * Extract file
     * @param string $filePath
     * @return array
     * @throws XlsxFileException
     * @throws XlsxParsingException
     * @returns array With "document" key, "dom" and "archive" key both are paths. "document" points to the xl/sharedStrings.xml
     * and "archive" points to the root of the archive. "dom" is the DOMDocument object for the slide.xml.
     */
    protected function prepareDocumentForReading(string $filePath): array
    {
        //Make sure we have a complete and correct path
        $filePath = realpath($filePath) ?: $filePath;

        $tempPath = $this->temporaryDirectory . DIRECTORY_SEPARATOR . uniqid();

        if (file_exists($tempPath)) {
            $this->rmdirRecursive($tempPath);
        }
        mkdir($tempPath);

        // Open the zip
        $this->openZip($filePath, $tempPath);

        // Find sharedStrings.xml
        $sharedStringPath = $tempPath . DIRECTORY_SEPARATOR . 'xl' . DIRECTORY_SEPARATOR . 'sharedStrings.xml';

        $documentXmlContents = file_get_contents($sharedStringPath);
        $dom = new DOMDocument();
        $loadXMLResult = $dom->loadXML($documentXmlContents, LIBXML_NOERROR | LIBXML_NOWARNING);

        if (!$loadXMLResult || !($dom instanceof DOMDocument)) {
            throw new XlsxParsingException( 'Could not parse XML document' );
        }

        return [
            "dom" => $dom,
            "document" => $sharedStringPath,
            "archive" => $tempPath
        ];
    }

    /**
     * @param DOMDocument $dom
     * @param string $archiveLocation
     * @param string $saveLocation
     * @throws XlsxFileException
     */
    protected function saveDocument(DOMDocument $dom, string $archiveLocation, string $saveLocation)
    {
        $document = $archiveLocation . DIRECTORY_SEPARATOR . 'xl' . DIRECTORY_SEPARATOR . 'sharedStrings.xml';

        if (!file_exists($document)) {
            throw new XlsxFileException( 'document.xml not found' );
        }

        $newDocumentXMLContents = $dom->saveXml();
        file_put_contents($document, $newDocumentXMLContents);

        $this->buildZip($saveLocation, $archiveLocation);
    }

    /**
     * Helper to remove tmp dir
     *
     * @param $dir
     * @return bool
     */
    protected function rmdirRecursive($dir)
    {
        $files = array_diff(scandir($dir), array('.', '..'));
        foreach($files as $file) {
            (is_dir("$dir/$file")) ? rmdirRecursive("$dir/$file") : unlink("$dir/$file");
        }
        return rmdir($dir);
    }

}
