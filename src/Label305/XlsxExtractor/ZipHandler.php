<?php

namespace Label305\XlsxExtractor;

use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use ZipArchive;

abstract class ZipHandler {

    /**
     * @param string $filePath
     * @param string $tempPath
     * @throws XlsxFileException
     */
    protected function openZip(string $filePath, string $tempPath)
    {
        $zip = new ZipArchive;
        $opened = $zip->open($filePath);
        if ($opened !== TRUE) {
            throw new XlsxFileException( 'Could not open zip archive ' . $filePath . '[' . $opened . ']' );
        }
        $zip->extractTo($tempPath);
        $zip->close();
    }

    /**
     * @param string $saveLocation
     * @param string $archiveLocation
     * @throws XlsxFileException
     */
    protected function buildZip(string $saveLocation, string $archiveLocation)
    {
        //Create a pptx file again
        $zip = new ZipArchive;

        $opened = $zip->open($saveLocation, ZIPARCHIVE::CREATE | ZipArchive::OVERWRITE);
        if ($opened !== true) {
            throw new XlsxFileException( 'Cannot open zip: ' . $saveLocation . ' [' . $opened . ']' );
        }

        // Create recursive directory iterator
        $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($archiveLocation), RecursiveIteratorIterator::LEAVES_ONLY);

        foreach($files as $name => $file) {
            $filePath = $file->getRealPath();
            if (in_array($file->getFilename(), array('.', '..'))) {
                continue;
            }
            if (!file_exists($filePath)) {
                throw new XlsxFileException( 'File does not exists: ' . $file->getPathname() );
            } else {
                if (!is_readable($filePath)) {
                    throw new XlsxFileException( 'File is not readable: ' . $file->getPathname() );
                } else {
                    if (!$zip->addFile($filePath, substr($file->getPathname(), strlen($archiveLocation) + 1))) {
                        throw new XlsxFileException( 'Error adding file: ' . $file->getPathname() );
                    }
                }
            }
        }
        if (!$zip->close()) {
            throw new XlsxFileException( 'Could not create zip file' );
        }
    }

}
