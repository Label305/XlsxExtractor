<?php

namespace Label305\XlsxExtractor\Decorated;

/**
 * Class Style
 * @package Label305\XlsxExtractor\Decorated
 *
 * Represents the style contents of a <w:rPr> object in the xlsx format.
 */
class Style {

    /**
     * @var string|null
     */
    public $rFont;
    /**
     * @var string|null
     */
    public $color;
    /**
     * @var string|null
     */
    public $family;
    /**
     * @var string|null
     */
    public $sz;
    /**
     * @var string|null
     */
    public $scheme;

    function __construct(?string $rFont = null, ?string $color = null, ?string $family = null, ?string $sz = null, ?string $scheme = null) {
        $this->rFont = $rFont;
        $this->color = $color;
        $this->family = $family;
        $this->sz = $sz;
        $this->scheme = $scheme;
    }

    /**
     * To docx xml string
     *
     * @return string
     */
    public function toXlsxXML()
    {
        $value = '';
        if ($this->rFont !== null) {
            $value .= '<rFont val="' . $this->rFont . '" />';
        }
        if ($this->color !== null) {
            $value .= '<color theme="' . $this->color . '"/>';
        }
        if ($this->family !== null) {
            $value .= '<family val="' . $this->family . '"/>';
        }
        if ($this->sz !== null) {
            $value .= '<sz val="' . $this->sz . '"/>';
        }
        if ($this->scheme !== null) {
            $value .= '<scheme val="' . $this->scheme . '"/>';
        }

        return $value;
    }
}