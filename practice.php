<?php

$textStorage = [];

function add(&$textStorage, $title, $text)
{
    $element = ['title' => $title, 'text' => $text];
    $textStorage[] = $element;
}

add($textStorage, 'Спорт', 'Плавание');
add($textStorage, 'Животные', 'Антилопа');

var_dump($textStorage);

function remove($i, &$textStorage)
{
    if (array_key_exists($i, $textStorage)) {
        unset($textStorage[$i]);
        return 'true';
    } else {
        return 'false';
    }
}

echo remove(0, $textStorage);
echo remove(5, $textStorage);
echo PHP_EOL;
echo PHP_EOL;

var_dump($textStorage);

function edit(int $i, string $title, string $text, &$textStorage)
{
    if (array_key_exists($i, $textStorage)) {
        $textStorage[$i]['title'] = $title;
        $textStorage[$i]['text'] = $text;
        return 'true';
    } else {
        return 'false';
    }
}

echo PHP_EOL;
echo edit(1, 'Животные', 'Бобёр', $textStorage);
echo PHP_EOL;
var_dump($textStorage);
