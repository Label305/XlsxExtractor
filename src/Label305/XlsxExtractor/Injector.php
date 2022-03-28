<?php namespace Label305\XlsxExtractor;


interface Injector {

    /**
     * @param array $mapping
     * @param string $fileToInjectLocationPath
     * @param string $saveLocationPath
     * @throws XlsxParsingException
     * @throws XlsxFileException
     * @return void
     */
    public function injectMappingAndCreateNewFile(array $mapping, string $fileToInjectLocationPath, string $saveLocationPath): void;

}