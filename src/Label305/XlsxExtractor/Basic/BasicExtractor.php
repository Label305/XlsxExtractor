<?php

namespace Label305\XlsxExtractor\Basic;


use DOMDocument;
use DOMNode;
use DOMText;
use Label305\XlsxExtractor\XlsxFileException;
use Label305\XlsxExtractor\XlsxHandler;
use Label305\XlsxExtractor\XlsxParsingException;
use Label305\XlsxExtractor\Extractor;


class BasicExtractor extends XlsxHandler implements Extractor {

    /**
     * @var int
     */
    protected $nextTagIdentifier;

    /**
     * @param $originalFilePath
     * @param $mappingFileSaveLocationPath
     * @return array
     * @throws XlsxParsingException
     * @throws XlsxFileException
     */
    public function extractStringsAndCreateMappingFile(string $originalFilePath, string $mappingFileSaveLocationPath): array
    {
        $prepared = $this->prepareDocumentForReading($originalFilePath);

        $this->nextTagIdentifier = 0;
        $result = $this->replaceAndMapValues($prepared['dom']->documentElement);

        $this->saveDocument($prepared['dom'], $prepared['archive'], $mappingFileSaveLocationPath);

        return $result;
    }

    /**
     * Override this method to make a more complex replace and mapping
     *
     * @param DOMNode $node
     * @return array returns the mapping array
     */
    protected function replaceAndMapValues(DOMNode $node): array
    {
        $result = [];

        if ($node instanceof DOMText) {
            $result[$this->nextTagIdentifier] = $node->nodeValue;
            $node->nodeValue = "%". $this->nextTagIdentifier. "%";
            $this->nextTagIdentifier++;
        }

        if ($node->childNodes !== null) {
            foreach ($node->childNodes as $child) {
                $result = array_merge(
                    $result,
                    $this->replaceAndMapValues($child)
                );
            }
        }

        return $result;
    }

}