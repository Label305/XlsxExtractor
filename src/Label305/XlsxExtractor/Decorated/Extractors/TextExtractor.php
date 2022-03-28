<?php

namespace Label305\XlsxExtractor\Decorated\Extractors;

use DOMElement;
use Label305\XlsxExtractor\Decorated\SharedStringPart;

interface TextExtractor {

    /**
     * @param DOMElement $DOMElement
     * The result is the array which contains te sentences
     * @return SharedStringPart[]|array
     */
    public function extract(DOMElement $DOMElement): array;
}