<?php

require 'vendor/autoload.php';


// config / params / const - все то, что необходимо будет вынести в "файлы хранения"
// new

function pToT($n){
    return \PhpOffice\PhpWord\Shared\Converter::pixelToTwip($n);
}

// госты
// значение одного сантиметра word в twip-ах.
define('CM', 37.75);

// разобраться с константным ассоциативным массивом
define('GOST_SECTION', array(
                    // ориентация: альбомная
                    'orientation' => 'portrait',
                    // поля страницы: вверхнее - 2 см, нижнее - 2 см, левое - 2 см, правое - 1 см
                    'marginTop' => pToT(2*CM),
                    'marginBottom' =>  pToT(2*CM),
                    'marginLeft' => pToT(2*CM),
                    'marginRight' =>  pToT(CM),
                    )
);

define('GOST_PARARAPH', array(
                    // междустрочный интервал: 1.5 строки
                    'space' => array('line' => 120),
                    'spaceBefore' => pToT(0),
                    'spaceAfter' => pToT(0),
                    )
);

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

function addFirstHeader($section, $text) {    
    //$textStyle = is_array($textStyle) ? $textStyle : array();
    $textStyle = array(
        'size' => 16,
        'bold' => TRUE,

    );
    //$paragraphStyle = is_array($paragraphStyle) ? $paragraphStyle : array();
    $paragraphStyle = array(
        // междустрочный интервал: 1.5 строки
        'space' => array('line' => 120),
        'spaceAfter' => pToT(32),
        // в ворде это 1.25, возможно найду более красивое число....
        'indentation' => array('firstLine' => pToT(47.1)),
        'alignment'=> 'both',
    );

    $section->addText(
        $text, 
        $textStyle,
        $paragraphStyle
        );
}

function addSimpleText($section, $text) {
    $textStyle = array(

    );

    $paragraphStyle = array(
        // междустрочный интервал: 1.5 строки
        'space' => array('line' => 120),
        // в ворде это 1.25, возможно найду более красивое число....
        'indentation' => array('firstLine' => pToT(47.1)),
        'spaceAfter' => 0,
        'alignment'=> 'both',
        
    );

    $section->addText(
        $text,
        $textStyle,
        $paragraphStyle
    );
}


function textParser($section, $text){

    $actionType = array(
        '#' => addFirstHeader,
        'text' => addSimpleText,

    );

    $text = explode('|n', $text);

    foreach($text as $s) {
        // первый символ в строке
        $fs = substr(htmlspecialchars($s), 0, 1);
        if (array_key_exists($fs,$actionType)) {

            // убираем спец. символ
            $s = str_replace('#', '', $s);

            $actionType[$fs]($section, htmlspecialchars(strval($s)));
        } else {
            $actionType['text']($section, htmlspecialchars(strval($s)));
        }
        //$actionType[ substr(htmlspecialchars($s), 0, 1) ]($section, htmlspecialchars(strval($s)), 'Default', $paragraphStyle);
        
    }
}



$phpWord = new \PhpOffice\PhpWord\PhpWord();

$phpWord->setDefaultFontName('Times New Roman');
$phpWord->setDefaultFontSize(14);

$properties = initDefaultDocSettings($phpWord);



$sectionStyle = GOST_SECTION;
$section = $phpWord->addSection($sectionStyle);

// print_r($_REQUEST[text]);
$text = $_REQUEST[text];

// $text = str_replace('|n', '\n', $text);
//$text = explode('|n', $text);
$paragraphStyle = GOST_PARARAPH;
textParser($section, $text);
//addFirstHeader($section, $text, 'Default', $paragraphStyle);
//$section->addText(
//    htmlspecialchars($text), 
//    $textStyle,
//    $paragraphStyle
//   );

$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter(
    $phpWord, 'Word2007'
);
$objWriter->save('docs/doc.docx');

echo 'all ok';