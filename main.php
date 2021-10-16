<?php

require 'vendor/autoload.php';


// config / params / const - все то, что необходимо будет вынести в "файлы хранения"
// new
function initDefaultDocSettings($phpWord) {
    $properties = $phpWord->getDocInfo();

    $properties->setCreator('Georgy');
    $properties->setCompany('Sibsiu');
    $properties->setTitle('My Title');
    $properties->getDescription('Generated doc file');
    $properties->setCategory('My category');
    $properties->setLastModifiedBy('My name');
    $properties->setCreated(mktime(0,0,0,10,16,2021));
    $properties->setModified(mktime(0,0,0,10,16,2021));
    $properties->setSubject('My Subject');
    $properties->setKeywords('my, key, word');
}

function addFirstHeader($section, $text, $textStyle = 'Default', $paragraphStyle = 'Default') {
    
    $textStyle = is_array($textStyle) ? $textStyle : array();
    $paragraphStyle = is_array($paragraphStyle) ? $paragraphStyle : array();
    
    $section->addText(
        htmlspecialchars($text), 
        $textStyle,
        $paragraphStyle
        );
}

$actionType = array(
    "#" => addFirstHeader,
);

$phpWord = new \PhpOffice\PhpWord\PhpWord();

$phpWord->setDefaultFontName('Times New Roman');
$phpWord->setDefaultFontSize(14);

$properties = initDefaultDocSettings($phpWord);

$sectionStyle = array(
    'orientation' => 'landscape',
    'marginBottom' =>  \PhpOffice\PhpWord\Shared\Converter::pixelToTwip(150)
);
$section = $phpWord->addSection($sectionStyle);

// print_r($_REQUEST[text]);
$text = $_REQUEST[text];

// $text = str_replace('|n', '\n', $text);
//$text = explode('|n', $text);

addFirstHeader($section, $text);
//$section->addText(
//    htmlspecialchars($text), 
//    $textStyle,
//    $paragraphStyle
//   );

$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter(
    $phpWord, 'Word2007'
);
$objWriter->save('docs/doc.docx');

