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
    public static function paragraphWithHTML(string $html,?SharedString $originalSharedString = null)
    {
        $html = "<html>" . strip_tags($html, '<br /><br><b><strong><em><i><u><font>') . "</html>";
        $html = str_replace("<br>", "<br />", $html);
        $html = str_replace("&nbsp;", " ", $html);
        $htmlDom = new DOMDocument;
        @$htmlDom->loadXml(preg_replace('/&(?!#?[a-z0-9]+;)/', '&amp;', html_entity_decode($html)));

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
     * @param bool $hasStyle
     */
    public function fillWithHTMLDom(
        DOMNode $node,
        ?SharedString $originalSharedString = null,
        $bold = false,
        $italic = false,
        $underline = false,
        $hasStyle = false
    ) {
        if ($node instanceof DOMText) {
            $originalStyle = null;
            if ($originalSharedString !== null) {
                $originalStyle = $this->getOriginalStyle($node, $originalSharedString);
            }
            $this[] = new SharedStringPart($node->nodeValue, $bold, $italic, $underline, $originalStyle);

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
                if ($node->nodeName == 'font') {
                    $hasStyle = true;
                }

                foreach ($node->childNodes as $child) {
                    $this->fillWithHTMLDom($child, $originalSharedString, $bold, $italic, $underline, $hasStyle);
                }
            }
        }
    }

    /**
     * @param DOMText $node
     * @param SharedString $originalSharedString
     * @return Style|null
     */
    private function getOriginalStyle(DOMText $node, SharedString $originalSharedString)
    {
        $originalStyle = null;
        if (array_key_exists($this->nextTagIdentifier, $originalSharedString)) {
            // Sometimes we extract a single space, but in the Paragraph the space is at the beginning of the sentence
            $startsWithSpace = strlen($node->nodeValue) > strlen(ltrim($node->nodeValue));
            if ($startsWithSpace && strlen(ltrim($originalSharedString[$this->nextTagIdentifier]->text)) === 0) {
                // When the current paragraph has no lengt it may be the space at the beginning
                $this->nextTagIdentifier++;
                // Return the next paragraph style
                if (array_key_exists($this->nextTagIdentifier, $originalSharedString)) {
                    $originalStyle = $originalSharedString[$this->nextTagIdentifier]->style;
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
    public function toHTML()
    {
        $result = '';

        $boldIsActive = false;
        $italicIsActive = false;
        $underlineIsActive = false;
        $styleActive = false;

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

            $openStyle = false;
            if ($sharedStringPart->style !== null && !$styleActive) {
                $styleActive = true;
                $openStyle = true;
            }

            $nextSharedStringPart = ($i + 1 < count($this)) ? $this[$i + 1] : null;
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

            $closeStyle = false;
            if ($nextSharedStringPart === null || ($nextSharedStringPart->style === null && $styleActive)) {
                $styleActive = false;
                $closeStyle = true;
            }

            $result .= $sharedStringPart->toHTML($openBold, $openItalic, $openUnderline, $openStyle, $closeBold, $closeItalic, $closeUnderline, $closeStyle);
        }

        return $result;
    }
}
