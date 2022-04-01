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
        $this->assertEquals("bold", $mapping[8][0]->text);
        $this->assertEquals(" and ", $mapping[9][0]->text);
        $this->assertEquals("colored", $mapping[10][0]->text);
        $this->assertEquals(" text", $mapping[11][0]->text);

        $mapping[0][0]->text = "Titel voor sheet";
        $mapping[7][0]->text = "Dit is een andere omschrijving met deels ";
        $mapping[8][0]->text = "dikgedrukt";
        $mapping[9][0]->text = " en ";
        $mapping[10][0]->text = "gekleurde";
        $mapping[11][0]->text = " tekst";

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__. '/fixtures/markup-extracted.xlsx', __DIR__. '/fixtures/markup-injected.xlsx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__. '/fixtures/markup-injected.xlsx', __DIR__. '/fixtures/markup-injected-extracted.xlsx');

        $this->assertEquals("Titel voor sheet", $otherMapping[0][0]->text);
        $this->assertEquals("Dit is een andere omschrijving met deels ", $otherMapping[7][0]->text);
        $this->assertEquals("dikgedrukt", $otherMapping[8][0]->text);
        $this->assertEquals(" en ", $otherMapping[9][0]->text);
        $this->assertEquals("gekleurde", $otherMapping[10][0]->text);
        $this->assertEquals(" tekst", $otherMapping[11][0]->text);

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
            0 => 'Titel voor slide',
            7 => 'Dit is een andere omschrijving met deels',
            8 => 'dikgedruk',
            9 => ' en ',
            10 => 'gekleurde',
            11 => ' tekst',
        ];

        foreach ($translations as $key => $translation) {
            $mapping[$key] = SharedString::paragraphWithHTML($translation, $mapping[$key]);
        }

        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, __DIR__. '/fixtures/markup-extracted.xlsx', __DIR__. '/fixtures/markup-injected.xlsx');

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile(__DIR__. '/fixtures/markup-injected.xlsx', __DIR__. '/fixtures/markup-injected-extracted.xlsx');

        $this->assertEquals('Titel voor slide', $otherMapping[0][0]->text);
        $this->assertEquals('Dit is een andere omschrijving met deels', $otherMapping[7][0]->text);
        $this->assertEquals('dikgedruk', $otherMapping[8][0]->text);
        $this->assertEquals(' en ', $otherMapping[9][0]->text);
        $this->assertEquals('gekleurde', $otherMapping[10][0]->text);
        $this->assertEquals(' tekst', $otherMapping[11][0]->text);

        unlink(__DIR__.'/fixtures/markup-extracted.xlsx');
        unlink(__DIR__.'/fixtures/markup-injected-extracted.xlsx');
        unlink(__DIR__.'/fixtures/markup-injected.xlsx');
    }

    /**
     * When a file contains special characters (i.e. `<`, `>`),
     * These should also be present in the extracted mapping
     */
    public function testSpecialCharactersInFile()
    {
        /* Given */
        $file = __DIR__ . '/fixtures/encoding.xlsx';
        $extractedFile = __DIR__ . '/fixtures/encoding-extracted.xlsx';

        /* When */
        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile($file, $extractedFile);

        /* Then */
        // text should contain encoded translations
        $this->assertEquals("3 < 5", $mapping[0][0]->text);
        $this->assertEquals("<font> tag is deprecated", $mapping[1][0]->text);
        $this->assertEquals(">< is not an X", $mapping[2][0]->text);

        unlink($extractedFile);
    }

    /**
     * When translations are injected with encoded characters (i.e. &lt;, &gt;),
     * These should also be present and encoded when extracting the injected file
     */
    public function testEncodedCharactersInTranslation()
    {
        /* Given */
        $file = __DIR__ . '/fixtures/encoding.xlsx';
        $extractedFile = __DIR__ . '/fixtures/encoding-extracted.xlsx';
        $injectedFile = __DIR__ . '/fixtures/encoding-injected.xlsx';
        $extractedInjectedFile = __DIR__ . '/fixtures/encoding-extracted-injected.xlsx';

        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile($file, $extractedFile);

        // decoded (loadXml) and encoded again (toHTML)
        $mapping[0][0]->text = SharedString::paragraphWithHTML("tree &lt; five")->toHTML();
        $mapping[1][0]->text = SharedString::paragraphWithHTML("La balise &lt;font&gt; est depreciee")->toHTML();
        $mapping[2][0]->text = SharedString::paragraphWithHTML("&gt;&lt; nest pas un X")->toHTML();

        /* When */
        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, $extractedFile, $injectedFile);

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile($injectedFile, $extractedInjectedFile);

        /* Then */
        // text should contain encoded translations
        $this->assertEquals("tree &lt; five", $otherMapping[0][0]->text);
        $this->assertEquals("La balise &lt;font&gt; est depreciee", $otherMapping[1][0]->text);
        $this->assertEquals("&gt;&lt; nest pas un X", $otherMapping[2][0]->text);

        unlink($extractedFile);
        unlink($injectedFile);
        unlink($extractedInjectedFile);
    }

    public function testMultiLineContent()
    {
        /* Given */
        $file = __DIR__ . '/fixtures/multi-line-content.xlsx';
        $extractedFile = __DIR__ . '/fixtures/multi-line-content-extracted.xlsx';
        $injectedFile = __DIR__ . '/fixtures/multi-line-content-injected.xlsx';
        $extractedInjectedFile = __DIR__ . '/fixtures/multi-line-content-extracted-injected.xlsx';

        $extractor = new DecoratedTextExtractor();
        $mapping = $extractor->extractStringsAndCreateMappingFile($file, $extractedFile);

        $this->assertEquals("Omschrijving", $mapping[0][0]->text);
        $this->assertEquals("Item 614:", $mapping[1][0]->text);
        $this->assertEquals("
Zijn je trampolineveren toe aan vervanging? Door jouw trampoline te voorzien van nieuwe Brand Name veren, springt je trampoline weer als nieuw!

De goudkleurige veren van Brand Name zijn conisch van vorm waardoor je geweldige sprongen kan maken en zacht kan landen. Ook zijn de veren dubbel gegalvaniseerd, zodat ze bestand zijn tegen roestvorming. 

Deze Brand Name veren hebben een lengte van 140 mm en zijn geschikt voor onderstaande ", $mapping[2][0]->text);


        $mapping[0][0]->text = SharedString::paragraphWithHTML("Description")->toHTML();
        $mapping[1][0]->text = SharedString::paragraphWithHTML("Unit 614:")->toHTML();
        $mapping[2][0]->text = SharedString::paragraphWithHTML("
Do your trampoline springs need to be replaced? By providing your trampoline with new Brand Name springs, your trampoline will jump like new!

Brand Name's gold-colored feathers are conical in shape, allowing you to make amazing jumps and land softly. The springs are also double galvanized, so that they are resistant to rusting.

These Brand Name springs have a length of 140 mm and are suitable for the following ")->toHTML();

        /* When */
        $injector = new DecoratedTextInjector();
        $injector->injectMappingAndCreateNewFile($mapping, $extractedFile, $injectedFile);

        $otherExtractor = new DecoratedTextExtractor();
        $otherMapping = $otherExtractor->extractStringsAndCreateMappingFile($injectedFile, $extractedInjectedFile);

        /* Then */
        // text should contain encoded translations
        $this->assertEquals("Description", $otherMapping[0][0]->text);
        $this->assertEquals("Unit 614:", $otherMapping[1][0]->text);
        $this->assertEquals("
Do your trampoline springs need to be replaced? By providing your trampoline with new Brand Name springs, your trampoline will jump like new!

Brand Name's gold-colored feathers are conical in shape, allowing you to make amazing jumps and land softly. The springs are also double galvanized, so that they are resistant to rusting.

These Brand Name springs have a length of 140 mm and are suitable for the following ", $otherMapping[2][0]->text);

        unlink($extractedFile);
        unlink($injectedFile);
        unlink($extractedInjectedFile);
    }
    
}