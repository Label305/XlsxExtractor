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
    private $rFont;
    /**
     * @var string|null
     */
    private $color;
    /**
     * @var string|null
     */
    private $family;
    /**
     * @var string|null
     */
    private $sz;
    /**
     * @var string|null
     */
    private $scheme;

    function __construct($rFont, $color, $family, $sz, $scheme) {
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

    /**
     * @return string|null
     */
    public function getRFont(): ?string
    {
        return $this->rFont;
    }

    /**
     * @return string|null
     */
    public function getColor(): ?string
    {
        return $this->color;
    }

    /**
     * @return string|null
     */
    public function getFamily(): ?string
    {
        return $this->family;
    }

    /**
     * @return string|null
     */
    public function getSz(): ?string
    {
        return $this->sz;
    }

    /**
     * @return string|null
     */
    public function getScheme(): ?string
    {
        return $this->scheme;
    }
}