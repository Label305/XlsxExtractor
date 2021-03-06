<?php

namespace Label305\XlsxExtractor\Decorated;


use DOMDocument;
use DOMElement;
use DOMNode;
use DOMText;
use Label305\XlsxExtractor\Decorated\Extractors\RNodeTextExtractor;
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

    protected function replaceAndMapValues(DOMNode $node) {
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
     * @param $result
     * @return array
     */
    protected function replaceAndMapValuesForParagraph(DOMNode $DOMNode, &$result)
    {
        $firstTextChild = null;
        $otherNodes = [];
        $parts = new SharedString();

        if ($DOMNode->childNodes !== null) {
            foreach ($DOMNode->childNodes as $DOMNodeChild) {

                if ($DOMNodeChild instanceof DOMElement && $DOMNodeChild->nodeName === "t") {
                    $parts[] = new SharedStringPart($DOMNodeChild->nodeValue);
                    $firstTextChild = $DOMNodeChild;

                } elseif ($DOMNodeChild instanceof DOMElement && $DOMNodeChild->nodeName === "r") {

                    // Parse results
                    $sharedStringParts = (new RNodeTextExtractor())->extract($DOMNodeChild);
                    if (count($sharedStringParts) !== 0) {
                        foreach ($sharedStringParts as $sharedStringPart) {
                            $parts[] = $sharedStringPart;
                        }
                        if ($firstTextChild === null) {
                            $firstTextChild = $DOMNodeChild;
                        } else {
                            $otherNodes[] = $DOMNodeChild;
                        }
                    }

                } elseif ($DOMNodeChild instanceof DOMElement) {
                    $this->replaceAndMapValuesForParagraph($DOMNodeChild, $result);
                }
            }

            if ($firstTextChild !== null) {
                $replacementNode = new DOMText();
                $replacementNode->nodeValue = "%" . $this->nextTagIdentifier . "%";
                $DOMNode->replaceChild($replacementNode, $firstTextChild);

                foreach ($otherNodes as $otherNode) {
                    $DOMNode->removeChild($otherNode);
                }

                $result[$this->nextTagIdentifier] = $parts;
                $this->nextTagIdentifier++;
            }
        }

        return $result;
    }
}
