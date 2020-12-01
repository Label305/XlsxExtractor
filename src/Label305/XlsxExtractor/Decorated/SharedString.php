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
     * Convenience constructor for the user of the API
     * Strings with <br> <b> <i> and <u> tags are supported.
     * @param $html string
     * @return SharedString
     */
    public static function paragraphWithHTML($html)
    {
        $html = "<html>" . strip_tags($html, '<br /><br><b><strong><em><i><u>') . "</html>";
        $html = str_replace("<br>", "<br />", $html);
        $html = str_replace("&nbsp;", " ", $html);
        $htmlDom = new DOMDocument;
        @$htmlDom->loadXml(preg_replace('/&(?!#?[a-z0-9]+;)/', '&amp;', html_entity_decode($html)));

        $sharedString = new SharedString();
        if ($htmlDom->documentElement !== null) {
            $sharedString->fillWithHTMLDom($htmlDom->documentElement);
        }

        return $sharedString;
    }

    /**
     * Recursive method to fill paragraph from HTML data
     *
     * @param DOMNode $node
     * @param int $br
     * @param bool $bold
     * @param bool $italic
     * @param bool $underline
     */
    public function fillWithHTMLDom(
        DOMNode $node,
        $bold = false,
        $italic = false,
        $underline = false
    ) {
        if ($node instanceof DOMText) {

            $this[] = new SharedStringPart($node->nodeValue, $bold, $italic, $underline);

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
                    $this->fillWithHTMLDom($child, $bold, $italic, $underline);
                }
            }
        }
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

            $result .= $sharedStringPart->toHTML($openBold, $openItalic, $openUnderline, $closeBold, $closeItalic, $closeUnderline);
        }

        return $result;
    }
}
