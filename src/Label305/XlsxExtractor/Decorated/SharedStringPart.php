<?php

namespace Label305\XlsxExtractor\Decorated;

/**
 * Class SharedStringPart
 * @package Label305\XlsxExtractor\Decorated
 *
 * Represents a <si> object in the xlsx format.
 */
class SharedStringPart {

    /**
     * @var string
     */
    public $text;

    /**
     * @var bool
     */
    public $bold;

    /**
     * @var bool
     */
    public $italic;

    /**
     * @var bool
     */
    public $underline;

    /**
     * @var Style|null
     */
    public $style;


    function __construct(
        string $text,
        bool $bold = false,
        bool $italic = false,
        bool $underline = false,
        ?Style $style = null
    ) {
        $this->text = $text;
        $this->bold = $bold;
        $this->italic = $italic;
        $this->underline = $underline;
        $this->style = $style;
    }

    /**
     * To docx xml string
     *
     * @return string
     */
    public function toXlsxXML(): string
    {
        $value = '<r>';

        if ($this->hasMarkup()) {

            $value .= '<rPr>';
            if ($this->bold) {
                $value .= '<b/>';
            }
            if ($this->italic) {
                $value .= '<i/>';
            }
            if ($this->underline) {
                $value .= '<u/>';
            }
            if ($this->style !== null) {
                $value .= $this->style->toXlsxXML();
            }
            $value .= '</rPr>';
            $value .= '<t xml:space="preserve">' . htmlentities($this->text, ENT_XML1) . "</t>";
        } else {
            $value .= '<t>' . htmlentities($this->text, ENT_XML1) . "</t>";
        }

        $value .= '</r>';

        return $value;
    }

    private function hasMarkup(): bool
    {
        if ($this->style !== null) {
            return true;
        }

        return $this->bold !== null || $this->italic !== null || $this->underline !== null;
    }

    /**
     * Convert to HTML
     *
     * To prevent duplicate tags (e.g. <strong) and allow wrapping you can use the parameters. If they are set to false
     * a tag will not be opened or closed.
     *
     * @param bool $firstWrappedInBold
     * @param bool $firstWrappedInItalic
     * @param bool $firstWrappedInUnderline
     * @param bool $firstWrappedInFont
     * @param bool $lastWrappedInBold
     * @param bool $lastWrappedInItalic
     * @param bool $lastWrappedInUnderline
     * @param bool $lastWrappedInFont
     * @return string HTML string
     */
    public function toHTML(
        $firstWrappedInBold = true,
        $firstWrappedInItalic = true,
        $firstWrappedInUnderline = true,
        $firstWrappedInFont = true,
        $lastWrappedInBold = true,
        $lastWrappedInItalic = true,
        $lastWrappedInUnderline = true,
        $lastWrappedInFont = true
    ): string {
        $value = '';

        if ($this->bold && $firstWrappedInBold) {
            $value .= "<strong>";
        }
        if ($this->italic && $firstWrappedInItalic) {
            $value .= "<em>";
        }
        if ($this->underline && $firstWrappedInUnderline) {
            $value .= "<u>";
        }
        if ($this->style !== null && !$this->style->isEmpty() && $firstWrappedInFont) {
            $value .= "<font>";
        }

        $value .= htmlentities($this->text);

        if ($this->style !== null && !$this->style->isEmpty() && $lastWrappedInFont) {
            $value .= "</font>";
        }
        if ($this->underline && $lastWrappedInUnderline) {
            $value .= "</u>";
        }
        if ($this->italic && $lastWrappedInItalic) {
            $value .= "</em>";
        }
        if ($this->bold && $lastWrappedInBold) {
            $value .= "</strong>";
        }

        return $value;
    }
}