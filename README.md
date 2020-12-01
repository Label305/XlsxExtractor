Xlsx Extractor [![Build Status](https://travis-ci.org/Label305/XlsxExtractor.svg)](https://travis-ci.org/Label305/XlsxExtractor)
=============

PHP library for extracting and replacing string data in .xlsx files. Pptx files are zip archives filled with XML documents and assets. Their format is described by [OOXML](http://nl.wikipedia.org/wiki/Office_Open_XML). This library only manipulates the `xl/sharedStrings.xml` file.

Composer installation
---

```json
"require": {
    "label305/xlsx-extractor": "0.1.*"
}
```

Basic usage
----

Import the basic classes.

```php
use Label305\XlsxExtractor\Basic\BasicExtractor;
use Label305\XlsxExtractor\Basic\BasicInjector;
```

First we need to extract all the contents from an existing `xlsx` file. This can be done using the `BasicExtractor`. Calling `extractStringsAndCreateMappingFile` will create a new file which name you pass in the second argument. This new file contains references so the library knows where to later inject the altered text back into.

```php
$extractor = new BasicExtractor();
$mapping = $extractor->extractStringsAndCreateMappingFile(
    'simple.xlsx',
    'simple-extracted.xlsx'
  );
```

Now that you have extracted contents you can inspect the content of the resulting `$mapping` array. And if you wish to change the content you can simply modify it. The array key maps to a symbol in the `simple-extracted.xlsx`.

```php
echo $mapping[0]; // Slide number one
```

Now after you changed your content, you can save it back to a new file. In this case that file is `simple-injected.xlsx`.

```php
$mapping[0] = "Slide number one";

$injector = new BasicInjector();
$injector->injectMappingAndCreateNewFile(
    $mapping,
    'simple-extracted.xlsx',
    'simple-injected.xlsx'
  );
```

License
---------
Copyright 2020 Label305 B.V.

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

[http://www.apache.org/licenses/LICENSE-2.0](http://www.apache.org/licenses/LICENSE-2.0)

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
