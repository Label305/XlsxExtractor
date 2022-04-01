<?php

namespace Label305\XlsxExtractor\Decorated;

use DOMElement;
use DOMNode;
use DOMText;
use Label305\XlsxExtractor\XlsxFileException;
use Label305\XlsxExtractor\XlsxHandler;
use Label305\XlsxExtractor\XlsxParsingException;
use Label305\XlsxExtractor\Extractor;


class DecoratedTextExtractor extends XlsxHandler implements Extractor {

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

    protected function replaceAndMapValues(DOMNode $node): array {
        $result = [];

        if ($node instanceof DOMElement && $node->nodeName == "si") {
            $this->replaceAndMapValuesForParagraph($node, $result);

        } else {
            if ($node->childNodes !== null) {
                foreach ($node->childNodes as $child) {
                    $result = array_merge(
                        $result,
                        $this->replaceAndMapValues($child)
                    );
                }
            }
        }

        return $result;
    }

    /**
     * @param DOMNode $DOMNode
     * @param array $result
     * @return array
     */
    protected function replaceAndMapValuesForParagraph(DOMNode $DOMNode, array &$result)
    {
        $foundDomNode = null;
        $foundSharedStringPart = null;
        $sharedString = new SharedString();
        if ($DOMNode->childNodes !== null) {

            // Each <r> tag contains one <t> tag with contents
            foreach ($DOMNode->childNodes as $key => $DOMNodeChild) {
                if ($DOMNodeChild instanceof DOMElement && $DOMNodeChild->nodeName === "t") {

                    $sharedString[] = new SharedStringPart($DOMNodeChild->nodeValue);

                    $replacementNode = new DOMText();
                    $replacementNode->nodeValue = "%" . $this->nextTagIdentifier . "%";
                    $DOMNode->replaceChild($replacementNode, $DOMNodeChild);

                    $result[$this->nextTagIdentifier] = $sharedString;
                    $this->nextTagIdentifier++;

                    continue;

                } elseif ($DOMNodeChild instanceof DOMElement) {
                    $this->replaceAndMapValuesForParagraph($DOMNodeChild, $result);
                }
            }
        }

        return $result;
    }
}
