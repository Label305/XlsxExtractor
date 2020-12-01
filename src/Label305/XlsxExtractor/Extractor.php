<?php

namespace Label305\XlsxExtractor;

interface Extractor {

    /**
     * @param string $originalFilePath
     * @param string $mappingFileSaveLocationPath
     * @throws XlsxParsingException
     * @throws XlsxFileException
     * @return array The mapping of all the strings
     */
    public function extractStringsAndCreateMappingFile(string $originalFilePath, string $mappingFileSaveLocationPath): array;

}