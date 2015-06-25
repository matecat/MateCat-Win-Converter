<?php

/**
 * This script simply sends all the files in the input dir to the converter,
 * receives the converted file and saves it in the output directory.
 */

require 'config.php';
require '../LegacyOfficeConverter.php';

ob_implicit_flush(1);

$converter = new LegacyOfficeConverter(CONVERTER_HOST, CONVERTER_PORT);

// Create the output dir if doesn't exist.
@mkdir(OUTPUT_FOLDER_PATH);

foreach (scandir(INPUT_FOLDER_PATH) as $fileName) {
    $filePath = INPUT_FOLDER_PATH . DIRECTORY_SEPARATOR . $fileName;
    if (!is_file($filePath)) continue;

    echo "converting $filePath<br>";

    try {
        $fileType = LegacyOfficeConverter::getFileType($filePath);
        $converted = $converter->convert($fileType, file_get_contents($filePath));
        echo 'converted!<br>';

        $convertedFilePath = OUTPUT_FOLDER_PATH . DIRECTORY_SEPARATOR . $fileName . 'x';
        file_put_contents($convertedFilePath, $converted);
        echo "converted file saved in $convertedFilePath<br>";

    } catch (Exception $e) {
        echo $e . '<br>';
    }
    echo '<br>';
}