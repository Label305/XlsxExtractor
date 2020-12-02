<?php

namespace Label305\XlsxExtractor\Decorated\Extractors;

use DOMElement;
use Label305\XlsxExtractor\Decorated\SharedStringPart;
use Label305\XlsxExtractor\Decorated\Style;

class RNodeTextExtractor implements TextExtractor{

    /**
     * @param DOMElement $DOMElement
     * The result is the array which contains te sentences
     * @return SharedStringPart[]
     */
    public function extract(DOMElement $DOMElement)
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