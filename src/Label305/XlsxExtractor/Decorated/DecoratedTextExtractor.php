<?php

namespace Label305\XlsxExtractor\Decorated;


use DOMDocument;
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
                    $sharedStringParts = $this->parseRNode($DOMNodeChild);
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

    /**
     * @param DOMElement $DOMElement
     * @return array
     */
    protected function parseRNode(DOMElement $DOMElement)
    {
        $bold = false;
        $italic = false;
        $underline = false;
        $text = null;
        $style = null;
        $result = [];

        foreach ($DOMElement->childNodes as $rChild) {
            if ($rChild instanceof DOMElement) {
                $this->parseChildRNode($rChild, $result, $bold, $italic, $underline, $text, $style);
            }
        }

        return $result;
    }


    /**
     * @param $rChild
     * @param array $result
     * @param bool $bold
     * @param bool $italic
     * @param bool $underline
     * @param string|null $text
     * @param Style|null $style
     */
    protected function parseChildRNode(
        DOMElement $rChild,
        array &$result,
        bool &$bold,
        bool &$italic,
        bool &$underline,
        ?string &$text,
        ?Style &$style
    ) {
        switch ($rChild->nodeName) {
            case "rPr" :
                $rFont = null;
                $color = null;
                $family = null;
                $sz = null;
                $scheme = null;
                $hasStyle = false;

                foreach ($rChild->childNodes as $propertyNode) {
                    if ($propertyNode instanceof DOMElement) {
                        $this->parseStyle($propertyNode,$rFont,$color,$family,$sz,$scheme, $hasStyle);
                        $this->parseFormatting($propertyNode,$bold,$italic,$underline);
                    }
                }
                if ($hasStyle) {
                    $style = new Style($rFont, $color, $family, $sz, $scheme);
                }
                break;

            case "t" :
                $text = $rChild->nodeValue;
                break;
        }

        if ($text !== null && strlen($text) !== 0) {
            $result[] = new SharedStringPart($text, $bold, $italic, $underline, $style);

            // Reset
            $style = null;
            $text = null;
        }
    }

    /**
     * @param DOMElement $propertyNode
     * @param string|null $rFont
     * @param string|null $color
     * @param string|null $family
     * @param string|null $sz
     * @param string|null $scheme
     * @param bool $hasStyle
     */
    private function parseStyle(
        DOMElement $propertyNode,
        ?string &$rFont,
        ?string &$color,
        ?string &$family,
        ?string &$sz,
        ?string &$scheme,
        bool &$hasStyle
    ) {
        if ($propertyNode->nodeName === "rFont") {
            $rFont = $propertyNode->getAttribute('val');
            $hasStyle = true;
        } elseif ($propertyNode->nodeName === "color") {
            $color = $propertyNode->getAttribute('theme');
            $hasStyle = true;
        } elseif ($propertyNode->nodeName === "family") {
            $family = $propertyNode->getAttribute('val');
            $hasStyle = true;
        } elseif ($propertyNode->nodeName === "sz") {
            $sz = $propertyNode->getAttribute('val');
            $hasStyle = true;
        } elseif ($propertyNode->nodeName === "scheme") {
            $scheme = $propertyNode->getAttribute('val');
            $hasStyle = true;
        }
    }

    /**
     * @param DOMElement $propertyNode
     * @param bool $bold
     * @param bool $italic
     * @param bool $underline
     */
    private function parseFormatting(
        DOMElement $propertyNode,
        bool &$bold,
        bool &$italic,
        bool &$underline
    ) {
        if ($propertyNode->nodeName == "b") {
            $bold = true;
        } elseif ($propertyNode->nodeName == "i") {
            $italic = true;
        } elseif ($propertyNode->nodeName == "u") {
            $underline = true;
        }
    }
}
