## Чтение данных из xlsx

Простая читалка xlsx файлов.
Структура файла:

| Наименование  | Категория         | Цена |
| ------------- |:-----------------:| ----:|
| Товар 1       | Компьютеры        | 1600 |
| Товар 2       | Компьютеры        | 1600 |

**то есть первая колонка всегда наименование поля**

```php
require_once dirname(__FILE__) . '/vendor/autoload.php';

use Excel\Xlsx;

$Reader = new Xlsx\ExcelReader();

$i = 0;
$fields = [];
$arrays = [];
$Reader->read(array(
    'file' => dirname(__FILE__) . '/import.xlsx',
), function ($Reader, $data) use (& $fields, & $i, & $arrays) {
    $i++;
    if ($i == 1) {
        // Сбор колонок для обработки
        foreach ($data as $key => $field) {
            $field = trim($field);
            $fields[$key] = $field;
        }
        return true;
    }

    $array = [];
    foreach ($data as $k => $value) {
        $array[$fields[$k]] = $value;
    }
    $arrays[] = $array;
    return true;
});

echo '<pre>';
print_r($arrays);
die;
```