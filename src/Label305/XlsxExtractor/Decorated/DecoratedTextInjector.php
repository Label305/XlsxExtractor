<?php

namespace Label305\XlsxExtractor\Decorated;

use DOMNode;
use DOMText;
use Label305\XlsxExtractor\Injector;
use Label305\XlsxExtractor\XlsxFileException;
use Label305\XlsxExtractor\XlsxHandler;
use Label305\XlsxExtractor\XlsxParsingException;

class DecoratedTextInjector extends XlsxHandler implements Injector {

    /**
     * @param array $mapping
     * @param string $fileToInjectLocationPath
     * @param string $saveLocationPath
     * @throws XlsxParsingException
     * @throws XlsxFileException
     * @return void
     */
    public function injectMappingAndCreateNewFile(array $mapping, string $fileToInjectLocationPath, string $saveLocationPath)
    {
        $prepared = $this->prepareDocumentForReading($fileToInjectLocationPath);

        $this->assignMappedValues($prepared['dom']->documentElement, $mapping);

        $this->saveDocument($prepared['dom'], $prepared["archive"], $saveLocationPath);
    }

    /**
     * @param DOMNode $node
     * @param array $mapping
     */
    protected function assignMappedValues(DOMNode $node, array $mapping)
    {
        if ($node instanceof DOMText) {
            $results = [];
            preg_match("/%[0-9]*%/", $node->nodeValue, $results);

            if (count($results) > 0) {
                $key = trim($results[0], '%');

                $parent = $node->parentNode;
                foreach ($mapping[$key] as $sharedStringPart) {
                    $fragment = $parent->ownerDocument->createDocumentFragment();
                    $fragment->appendXML($sharedStringPart->toXlsxXML());
                    $parent->insertBefore($fragment, $node);
                }
                $parent->removeChild($node);
            }
        }

        if ($node->childNodes !== null) {
            foreach ($node->childNodes as $child) {
                $this->assignMappedValues($child, $mapping);
            }
        }
    }
}