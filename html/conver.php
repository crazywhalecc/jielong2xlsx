<?php

require_once "vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (!isset($_POST['filecontent'], $_POST['area_type'])) {
    echo '
<html>
<head>
<title>接龙转excel</title>
<meta charset="utf-8">
</head>
<body>
<form enctype="multipart/form-data" action="conver.php" method="post">
<textarea rows="10" cols="20" name="filecontent" type="text"></textarea>
<select name="area_type">
  <option value="qi">期</option>
  <option value="nong">弄</option>
</select>
<input type="submit" value="提交"/>
</form>
<pre>
将微信接龙的内容粘贴到上方框框中，可以将统计的楼号户号，信息摘出来变成excel。

注：请不要带有emoji表情等内容，最好是纯文本，否则转出来的excel可能无法打开。

例子：
1. 一期3号806酒精6瓶，防护服3套，隔离防护面罩6个.
2. 14–704 酒精5瓶
3. 一期12号1307 酒精10瓶，防护服10套，面罩10个，手套3盒
4. 一期3号1401，酒精1瓶
5. 一期1号楼2012室 酒精5瓶，手套一盒
7. 2期26-701 酒精4瓶
8. 1期-12-206  酒精5瓶
9. 1期-12-206  酒精5瓶 面罩2个 手套1盒
10. 一期14号604室梅雪 酒精1瓶
11. 二期28楼1706 酒精5瓶
12. 一期19号601 酒精8瓶 防护服3套
13. 1期-3号1403 酒精1瓶
14. 1期11号楼101室 酒精3瓶，手套1盒
15. 1期11-101酒精3瓶，手套1盒
</pre>
</body>
</html>
';
    exit();
}
$file = $_POST['filecontent'];
$type = ['qi' => '期', 'nong' => '弄'];

function filterEmoji($str)

{

    $str = preg_replace_callback(

        '/./u',

        function (array $match) {

            return strlen($match[0]) >= 4 ? '' : $match[0];
        },

        $str
    );

    return $str;
}
$file = filterEmoji($file);
//$file = file_get_contents("jielong.txt");
$file = str_replace("\r\n", "\n", $file);
$lines = explode("\n", $file);
global $result, $ls, $false, $exception, $type;
global $argv;
if (!isset($type[$_POST['area_type']])) {
    die('INVALID area_type INPUT: ' . $_POST['area_type']);
}

$type = $type[$_POST['area_type']];
array_shift($argv);
$result = [];
$ls = [];
$false = [];
$exception = [];
foreach ($lines as $line) {
    $line = explode('. ', $line);
    $num = $line[0];
    $content = $line[1];
    $content = replace_num($content);
    compile_full_text($content, $num);
}
resolve_no_qi();

ksort($result);



$spreadsheet = new Spreadsheet();
$worksheet = $spreadsheet->getActiveSheet();
$errorsheet = new Worksheet($spreadsheet, '有问题的');

//设置工作表标题名称
$worksheet->setTitle('统计表');

$worksheet->setCellValueByColumnAndRow(1, 1, '序号');
$worksheet->setCellValueByColumnAndRow(2, 1, $type);
$worksheet->setCellValueByColumnAndRow(3, 1, '楼栋');
$worksheet->setCellValueByColumnAndRow(4, 1, '室号');
$worksheet->setCellValueByColumnAndRow(5, 1, '附加信息');
$n = 6;
foreach ($argv as $v) {
    $worksheet->setCellValueByColumnAndRow($n, 1, '物品【' . $v . '】');
    ++$n;
}

$errorsheet->setCellValueByColumnAndRow(1, 1, '接龙序号');
$errorsheet->setCellValueByColumnAndRow(2, 1, '内容');

$ni = 2;
foreach ($result as $k => $v) {
    $worksheet->setCellValueByColumnAndRow(1, $ni, $k);
    $worksheet->setCellValueByColumnAndRow(2, $ni, $v[0]);
    $worksheet->setCellValueByColumnAndRow(3, $ni, $v[1]);
    $worksheet->setCellValueByColumnAndRow(4, $ni, $v[2]);
    $worksheet->setCellValueByColumnAndRow(5, $ni, $v[3]);
    $n = 6;
    foreach ($argv as $vs) {
        if (isset($v[4][$vs])) {
            $worksheet->setCellValueByColumnAndRow($n, $ni, $v[4][$vs]);
        } else {
            $worksheet->setCellValueByColumnAndRow($n, $ni, '0');
        }
        ++$n;
    }
    ++$ni;
}
$ni = 2;
foreach ($exception as $k => $v) {
    $errorsheet->setCellValueByColumnAndRow(1, $ni, $k);
    $errorsheet->setCellValueByColumnAndRow(2, $ni, $v);
    ++$ni;
}
$styleArrayBody = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
            'color' => ['argb' => '666666'],
        ],
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
];
$total_rows = count($result) + 2;
//添加所有边框/居中
$worksheet->getStyle('A1:E' . $total_rows)->applyFromArray($styleArrayBody);
$worksheet->getColumnDimension('B')->setAutoSize(true);
$worksheet->getColumnDimension('E')->setAutoSize(true);
$errorsheet->getColumnDimension('B')->setAutoSize(true);
$spreadsheet->addSheet($errorsheet, 1);
header('Content-Type: application/vnd.ms-excel'); // 设置文件类型
header('Content-Disposition: attachment;filename="' . 'output.xlsx"'); // 设置文件名称及文件后缀，文件名可使用变量替代。
header('Cache-Control: max-age=0'); // 设置缓存时间，此处为0不缓存
$writer = new Xlsx($spreadsheet);

$writer->save('php://output');

//$writer->save('world.xlsx');

function full_trim($txt)
{
    return trim($txt, "- \n\t，、,。");
}

function compile_full_text($content, $num)
{
    global $result, $ls, $false, $type;
    $patterns = [
        "*$type*号楼*室*",
        "*$type*号楼*",
        "*$type*楼*室*",
        "*$type*楼*",
        "*$type*号*室*",
        "*$type*号*",
        "*$type*栋*室*",
        "*$type*栋*",
        "*$type-*-*",
        "*$type*-*",
    ];
    foreach ($patterns as $id => $pattern) {
        if (($arg = match_args($pattern, $content)) !== false) {
            if (count($arg) < 3) die(json_encode($arg, 128 | 256));
            $r = match_room($arg, in_array($id, [0, 2, 4, 6]), "match[$id]:" . implode(",", $arg) . "");
            if (in_array($id, [0, 2, 4, 6])) {
                $content = $arg[3];
            } else {
                $content = str_replace($r[2], '', $arg[2]);
            }
            $content = str_replace(["，", "、", "；"], " ", full_trim($content));
            $r[] = $content;
            $r = array_merge($r, explode_items($content));
            $result[$num] = $r;
            //echo $num.PHP_EOL;
            $ls[$num] = $arg;
            break;
        }
    }
    if (!isset($ls[$num])) {
        $false[$num] = $content;
    }
}

function explode_items($content)
{
    global $argv;
    $q = [];
    foreach ($argv as $v) {
        if (($pos = mb_strpos($content, $v)) !== false) {
            $q[$v] = get_start_num(full_trim(mb_substr($content, $pos + mb_strlen($v))), 'Match:' . $content . ",$v");
        }
    }
    return [$q];
}

function match_room($arg, $end_room = false, $content = '')
{
    $qi = get_end_num($arg[0], $content);
    //if ($qi !== 1 && $qi !== 2) die('出错啦：' . json_encode($arg, 128 | 256));
    $hao = get_end_num($arg[1], $content);
    $room = $end_room ? get_end_num($arg[2], $content) : get_start_num($arg[2], $content);
    return [$qi, $hao, $room];
}

function resolve_no_qi()
{
    global $result, $false, $exception, $type;
    foreach ($false as $k => $v) {
        $patterns = [
            "*号楼*室*",
            "*号楼*",
            "*楼*室*",
            "*楼*",
            "*号*室*",
            "*号*",
            "*栋*室*",
            "*栋*",
            "*-*",
        ];
        foreach ($patterns as $id => $pattern) {
            if (($arg = match_args($pattern, $v)) !== false) {
                $hao = get_end_num($arg[0], $v);
                if ($type === 'qi')
                    array_unshift($arg, $hao >= 25 ? '2' : '1');
                else
                    array_unshift($arg, '0');
                $r = match_room($arg, in_array($id, [0, 2, 4, 6]));
                if (in_array($id, [0, 2, 4, 6])) {
                    $content = mb_substr($arg[3], mb_strpos($arg[3], "室"));
                } else {
                    $content = str_replace($r[2], '', $arg[2]);
                }
                if (mb_strpos($v, "荷塘月色") !== false) {
                    var_dump($content, $arg);
                }
                $r[] = full_trim($content);
                $result[$k] = $r;
                continue 2;
            }
        }
        if (!isset($result[$k])) {
            $exception[$k] = $v;
        }
    }
}

function get_end_num($str, $content = '')
{
    $str = trim($str, "-# ，,号");
    $n = '';
    $i = mb_strlen($str) - 1;
    while ($i >= 0 && is_numeric($p = mb_substr($str, $i, 1))) {
        $n = $p . $n;
        --$i;
    }
    if (!is_numeric($n)) die("数字解析错误end:" . $str . "\n$content");
    return (int) $n;
}

function get_start_num($str, $content = '')
{
    $str = trim($str,  "-# ，,号");
    $n = '';
    $i = 0;
    while ($i < mb_strlen($str) && is_numeric($p = mb_substr($str, $i, 1))) {
        $n .= strval($p);
        ++$i;
    }
    if (!is_numeric($n)) die("数字解析错误start:" . $str . "\n$content");
    return (int) $n;
}

function replace_num($str): string
{
    $str = str_replace(['一', '两', '二', '三', '四', '五', '六', '七', '八', '九', '十'], [' 1', ' 2', ' 2', ' 3', ' 4', ' 5', ' 6', ' 7', ' 8', ' 9'], $str);
    $str = str_replace(['瓶', '个', '箱', '盒', '篮', '斤', '袋', '幅', '蓝'], '', $str);
    for ($i = 0; $i < mb_strlen($str); ++$i) {
        if (mb_substr($str, $i, 1) === "套" && is_numeric(mb_substr($str, $i - 1, 1))) {
            $str = mb_substr($str, 0, $i) . mb_substr($str, $i + 1);
        }
    }
    $str = str_replace('—', '-', $str);
    $str = str_replace('–', '-', $str);
    $str = str_replace('－', '-', $str);
    $str = str_replace('～', '-', $str);
    $str = str_replace('#', '-', $str);
    return $str;
}

function match_args(string $pattern, string $subject)
{
    //*号*室*
    $result = [];
    if (match_pattern($pattern, $subject)) {
        if (mb_strpos($pattern, '*') === false) {
            return [];
        }
        $exp = explode('*', $pattern);
        $i = 0;
        foreach ($exp as $k => $v) {
            if (empty($v) && $v !== "0" && $k === 0) {
                continue;
            }
            if (empty($v) && $v !== "0" && $k === count($exp) - 1) {
                $subject .= '^EOL';
                $v = '^EOL';
            }
            $cur_var = '';
            $ori = $i;
            while (($a = mb_substr($subject, $i, mb_strlen($v))) !== $v && !(empty($a) && $a !== "0")) {
                $cur_var .= mb_substr($subject, $i, 1);
                ++$i;
            }
            if ($i !== $ori || $k === 1 || $k === count($exp) - 1) {
                $result[] = $cur_var;
            }
            $i += mb_strlen($v);
        }
        return $result;
    }
    return false;
}

function match_pattern(string $pattern, string $subject): bool
{
    if (empty($pattern) && empty($subject) && $subject !== "0") {
        return true;
    }
    if (mb_strpos($pattern, '*') === 0 && mb_substr($pattern, 1, 1) !== '' && empty($subject) && $subject !== "0") {
        return false;
    }
    if (mb_strpos($pattern, mb_substr($subject, 0, 1)) === 0) {
        return match_pattern(mb_substr($pattern, 1), mb_substr($subject, 1));
    }
    if (mb_strpos($pattern, '*') === 0) {
        return match_pattern(mb_substr($pattern, 1), $subject) || match_pattern($pattern, mb_substr($subject, 1));
    }
    return false;
}
