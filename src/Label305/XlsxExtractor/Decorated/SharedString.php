<?php namespace Label305\XlsxExtractor\Decorated;

use ArrayObject;
use DOMDocument;
use DOMNode;
use DOMText;

/**
 * Class Paragraph
 * @package Label305\XlsxExtractor\Decorated
 *
 * Represents a list of <si> objects in the xlsx format. Does not contain
 * <w:p> data. That data is preserved in the extracted document.
 */
class SharedString extends ArrayObject
{
    /**
     * @var int
     */
    protected $nextTagIdentifier = 0;

    /**
     * Convenience constructor for the user of the API
     * Strings with <br> <b> <i> <u> and <font> tags are supported.
     * @param $html string
     * @param SharedString|null $originalSharedString
     * @return SharedString
     */
    public static function paragraphWithHTML(string $html,?SharedString $originalSharedString = null): SharedString
    {
        $html = "<html>" . $html . "</html>";
        $html = str_replace("<br>", "<br />", $html);
        $html = str_replace("&nbsp;", " ", $html);
        $htmlDom = new DOMDocument;
        @$htmlDom->loadXml(preg_replace('/&(?!#?[a-z0-9]+;)/', '&amp;', $html));

        $sharedString = new SharedString();
        if ($htmlDom->documentElement !== null) {
            $sharedString->fillWithHTMLDom($htmlDom->documentElement, $originalSharedString);
        }

        return $sharedString;
    }

    /**
     * Recursive method to fill paragraph from HTML data
     *
     * @param DOMNode $node
     * @param SharedString|null $originalSharedString
     * @param bool $bold
     * @param bool $italic
     * @param bool $underline
     */
    public function fillWithHTMLDom(
        DOMNode $node,
        ?SharedString $originalSharedString = null,
        bool $bold = false,
        bool $italic = false,
        bool $underline = false
    ) {
        if ($node instanceof DOMNode && ($node->nodeName === 'a' || $node->nodeName === 'span') || $node->nodeName === 'div') {
            $this[] = new SharedStringPart($node->ownerDocument->saveXML($node), $bold, $italic, $underline, null);
            $this->nextTagIdentifier++;

        } else if ($node instanceof DOMText) {
            $originalStyle = null;
            if ($originalSharedString !== null) {
                $originalStyle = $this->getOriginalStyle($node, $originalSharedString);
            }
            $this[] = new SharedStringPart($node->nodeValue, $bold, $italic, $underline, $originalStyle);
            $this->nextTagIdentifier++;

        } else {
            if ($node->childNodes !== null) {

                if ($node->nodeName == 'b' || $node->nodeName == 'strong') {
                    $bold = true;
                }
                if ($node->nodeName == 'i' || $node->nodeName == 'em') {
                    $italic = true;
                }
                if ($node->nodeName == 'u') {
                    $underline = true;
                }

                foreach ($node->childNodes as $child) {
                    $this->fillWithHTMLDom($child, $originalSharedString, $bold, $italic, $underline);
                }
            }
        }
    }

    /**
     * @param DOMText $node
     * @param SharedString $originalSharedString
     * @return Style|null
     */
    private function getOriginalStyle(DOMText $node, SharedString $originalSharedString): ?Style
    {
        $originalStyle = null;
        if (array_key_exists($this->nextTagIdentifier, $originalSharedString->getArrayCopy())) {
            // Sometimes we extract a single space, but in the Paragraph the space is at the beginning of the sentence
            $startsWithSpace = strlen($node->nodeValue) > strlen(ltrim($node->nodeValue));
            if ($startsWithSpace && strlen(ltrim($originalSharedString[$this->nextTagIdentifier]->text)) === 0) {
                // When the current paragraph has no length it may be the space at the beginning
                if (array_key_exists($this->nextTagIdentifier + 1, $originalSharedString->getArrayCopy())) {
                    // Add the next paragraph style
                    $originalStyle = $originalSharedString[$this->nextTagIdentifier + 1]->style;
                    $this->nextTagIdentifier++;
                }
            } else {
                $originalStyle = $originalSharedString[$this->nextTagIdentifier]->style;
            }
        }
        return $originalStyle;
    }

    /**
     * Give me a paragraph HTML
     *
     * @return string
     */
    public function toHTML(): string
    {
        $result = '';

        $boldIsActive = false;
        $italicIsActive = false;
        $underlineIsActive = false;
        $fontIsActive = false;

        for ($i = 0; $i < count($this); $i++) {

            $sharedStringPart = $this[$i];

            $openBold = false;
            if ($sharedStringPart->bold && !$boldIsActive) {
                $boldIsActive = true;
                $openBold = true;
            }

            $openItalic = false;
            if ($sharedStringPart->italic && !$italicIsActive) {
                $italicIsActive = true;
                $openItalic = true;
            }

            $openUnderline = false;
            if ($sharedStringPart->underline && !$underlineIsActive) {
                $underlineIsActive = true;
                $openUnderline = true;
            }

            $openFont = false;
            if ($sharedStringPart->style !== null && !$sharedStringPart->style->isEmpty() &&
                !$fontIsActive &&
                !$boldIsActive &&
                !$italicIsActive &&
                !$underlineIsActive &&
                count($this) > 1
            ) {
                $openFont = true;
                $fontIsActive = true;
            }

            $nextSharedStringPart = ($i + 1 < count($this)) ? $this[$i + 1] : null;

            $closeFont = false;
            if ($fontIsActive) {
                $closeFont = true;
                $fontIsActive = false;
            }

            $closeBold = false;
            if ($nextSharedStringPart === null || (!$nextSharedStringPart->bold && $boldIsActive)) {
                $boldIsActive = false;
                $closeBold = true;
            }

            $closeItalic = false;
            if ($nextSharedStringPart === null || (!$nextSharedStringPart->italic && $italicIsActive)) {
                $italicIsActive = false;
                $closeItalic = true;
            }

            $closeUnderline = false;
            if ($nextSharedStringPart === null || (!$nextSharedStringPart->underline && $underlineIsActive)) {
                $underlineIsActive = false;
                $closeUnderline = true;
            }

            $result .= $sharedStringPart->toHTML($openBold, $openItalic, $openUnderline, $openFont, $closeBold, $closeItalic, $closeUnderline, $closeFont);
        }

        return $result;
    }
}
