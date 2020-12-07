<?php

use Label305\XlsxExtractor\Basic\BasicExtractor;
use Label305\XlsxExtractor\Basic\BasicInjector;
use Label305\XlsxExtractor\Decorated\DecoratedTextExtractor;
use Label305\XlsxExtractor\Decorated\DecoratedTextInjector;
use Label305\XlsxExtractor\Decorated\SharedString;
use Label305\XlsxExtractor\Decorated\SharedStringPart;
use Label305\XlsxExtractor\Decorated\Style;

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

    public function test_markup() {

        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__.'/fixtures/markup.xlsx', __DIR__.'/fixtures/markup-extracted.xlsx');

        $this->assertEquals("Title for sheet", $mapping[0][0]->text);
        $this->assertEquals("This is another description with parially ", $mapping[7][0]->text);
        $this->assertEquals("bold", $mapping[7][1]->text);
        $this->assertEquals(" and ", $mapping[7][2]->text);
        $this->assertEquals("colored", $mapping[7][3]->text);
        $this->assertEquals(" text", $mapping[7][4]->text);
        $this->assertEquals('This is another description with parially <strong>bold</strong><font> and </font><strong>colored</font></strong> text</font>', $mapping[7]->toHTML());

        $mapping[0][0]->text = "Titel voor sheet";
        $mapping[7][0]->text = "Dit is een andere omschrijving met deels ";
        $mapping[7][1]->text = "dikgedrukt";
        $mapping[7][2]->text = " en ";
        $mapping[7][3]->text = "gekleurde";
        $mapping[7][4]->text = " tekst";

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__. '/fixtures/markup-extracted.xlsx', __DIR__. '/fixtures/markup-injected.xlsx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__. '/fixtures/markup-injected.xlsx', __DIR__. '/fixtures/markup-injected-extracted.xlsx');

        $this->assertEquals("Titel voor sheet", $otherMapping[0][0]->text);
        $this->assertEquals("Dit is een andere omschrijving met deels ", $otherMapping[7][0]->text);
        $this->assertEquals("dikgedrukt", $otherMapping[7][1]->text);
        $this->assertEquals(" en ", $otherMapping[7][2]->text);
        $this->assertEquals("gekleurde", $otherMapping[7][3]->text);
        $this->assertEquals(" tekst", $otherMapping[7][4]->text);

        unlink(__DIR__.'/fixtures/markup-extracted.xlsx');
        unlink(__DIR__.'/fixtures/markup-injected-extracted.xlsx');
        unlink(__DIR__.'/fixtures/markup-injected.xlsx');
    }

    public function test_sharedString_toHtml()
    {
        $sharedString = new SharedString();
        $sharedString[] = new SharedStringPart('This is a test with ');
        $sharedString[] = new SharedStringPart('bold' , true);
        $sharedString[] = new SharedStringPart(' and ');
        $sharedString[] = new SharedStringPart('italic' , false, true);
        $sharedString[] = new SharedStringPart(' and ');
        $sharedString[] = new SharedStringPart('underline' , false, false, true);
        $sharedString[] = new SharedStringPart(' and ');
        $sharedString[] = new SharedStringPart('colored text' , false, false, false, new Style());

        $this->assertEquals('This is a test with <strong>bold</strong> and <em>italic</em> and <u>underline</u> and colored text', $sharedString->toHTML());
    }

    public function test_sharedString_fillWithHTMLDom()
    {
        $html = 'This is a test with <strong>bold</strong> and <em>italic</em> and <u>underline</u> and <font>colored text</font>';
        $html = "<html>" . $html . "</html>";

        $htmlDom = new DOMDocument;
        @$htmlDom->loadXml($html);

        $sharedString = new SharedString();
        $sharedString->fillWithHTMLDom($htmlDom->documentElement);

        $this->assertEquals('This is a test with ', $sharedString[0]->text);
        $this->assertEquals('bold', $sharedString[1]->text);
        $this->assertTrue($sharedString[1]->bold);
        $this->assertEquals(' and ', $sharedString[2]->text);
        $this->assertEquals('italic', $sharedString[3]->text);
        $this->assertTrue($sharedString[3]->italic);
        $this->assertEquals(' and ', $sharedString[4]->text);
        $this->assertEquals('underline', $sharedString[5]->text);
        $this->assertTrue($sharedString[5]->underline);
        $this->assertEquals(' and ', $sharedString[6]->text);
        $this->assertEquals('colored text', $sharedString[7]->text);
    }

    public function test_paragraphWithHTML()
    {
        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile(__DIR__. '/fixtures/markup.xlsx', __DIR__. '/fixtures/markup-extracted.xlsx');

        $translations = [
            'Titel voor slide',
            'Dit is een andere omschrijving met deels <strong>dikgedruk</strong> en <strong>gekleurde</strong> <font>tekst</font>',
        ];

        foreach ($translations as $key => $translation) {
            $mapping[$key] = SharedString::paragraphWithHTML($translation, $mapping[$key]);
        }

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__. '/fixtures/markup-extracted.xlsx', __DIR__. '/fixtures/markup-injected.xlsx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__. '/fixtures/markup-injected.xlsx', __DIR__. '/fixtures/markup-injected-extracted.xlsx');

        $this->assertEquals('Titel voor slide', $otherMapping[0][0]->text);
        $this->assertEquals('Dit is een andere omschrijving met deels ', $otherMapping[1][0]->text);
        $this->assertEquals('dikgedruk', $otherMapping[1][1]->text);
        $this->assertEquals(' en ', $otherMapping[1][2]->text);
        $this->assertEquals('gekleurde', $otherMapping[1][3]->text);
        $this->assertEquals(' ', $otherMapping[1][4]->text);
        $this->assertEquals('tekst', $otherMapping[1][5]->text);

        unlink(__DIR__.'/fixtures/markup-extracted.xlsx');
        unlink(__DIR__.'/fixtures/markup-injected-extracted.xlsx');
        unlink(__DIR__.'/fixtures/markup-injected.xlsx');
    }
    
}