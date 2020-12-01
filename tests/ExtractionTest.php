<?php

use Label305\XlsxExtractor\Basic\BasicExtractor;
use Label305\XlsxExtractor\Basic\BasicInjector;

class ExtractionTest extends TestCase {

    public function test_multipleSlides() {

        $extractor = new BasicExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/simple.xlsx', __DIR__.'/fixtures/simple-extracted.xlsx');

        $this->assertEquals("Content 1 on sheet 1", $mapping[0]);
        $this->assertEquals("Content 2 on sheet 1", $mapping[1]);
        $this->assertEquals("Content 3 on sheet 1", $mapping[2]);
        $this->assertEquals("Content on sheet 2", $mapping[3]);

        $mapping[0] = "Inhoud 1 op sheet 1";
        $mapping[3] = "Inhoud op sheet 2";

        $injector = new BasicInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__. '/fixtures/simple-extracted.xlsx', __DIR__. '/fixtures/simple-injected.xlsx');

        $otherExtractor = new BasicExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__. '/fixtures/simple-injected.xlsx', __DIR__. '/fixtures/simple-injected-extracted.xlsx');

        $this->assertEquals("Inhoud 1 op sheet 1", $otherMapping[0]);
        $this->assertEquals("Inhoud op sheet 2", $otherMapping[3]);

        unlink(__DIR__.'/fixtures/simple-extracted.xlsx');
        unlink(__DIR__.'/fixtures/simple-injected-extracted.xlsx');
        unlink(__DIR__.'/fixtures/simple-injected.xlsx');
    }

    public function test_duplicateText() {

        $extractor = new BasicExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/simple-duplicate-text.xlsx', __DIR__.'/fixtures/simple-duplicate-text-extracted.xlsx');

        $this->assertEquals("Duplicate text", $mapping[4]);

        $mapping[4] = "Duplicaat tekst";

        $injector = new BasicInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__. '/fixtures/simple-duplicate-text-extracted.xlsx', __DIR__. '/fixtures/simple-duplicate-text-injected.xlsx');

        $otherExtractor = new BasicExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__. '/fixtures/simple-duplicate-text-injected.xlsx', __DIR__. '/fixtures/simple-duplicate-text-injected-extracted.xlsx');

        $this->assertEquals("Duplicaat tekst", $otherMapping[4]);

        unlink(__DIR__.'/fixtures/simple-duplicate-text-extracted.xlsx');
        unlink(__DIR__.'/fixtures/simple-duplicate-text-injected-extracted.xlsx');
        unlink(__DIR__.'/fixtures/simple-duplicate-text-injected.xlsx');
    }
    
}