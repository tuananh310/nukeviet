<?php

/**
 * @Project NUKEVIET 4.x
 * @Author VINADES.,JSC (contact@vinades.vn)
 * @Copyright (C) 2014 VINADES.,JSC. All rights reserved
 * @License GNU/GPL version 2 or any later version
 * @Createdate 3/9/2010 23:25
 */

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (!defined('NV_IS_FILE_ADMIN')) {
    die('Stop!!!');
}

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// Tạo tiêu đề
$sheet
    ->setCellValue('A1', 'STT')
    ->setCellValue('B1', 'Tên bài đăng')
    ->setCellValue('C1', 'Tác giả')
    ->setCellValue('D1', 'Ngày xuất bản')
    ->setCellValue('E1', 'Lượt xem')
    ->setCellValue('F1', 'Điểm đánh giá');

// Ghi dữ liệu
$db_slave->select('total_rating, title, publtime, hitstotal, author')->from(NV_PREFIXLANG . '_' . 'news_rows');
$result = $db_slave->query($db_slave->sql());
$rowCount = 2;

while (list($total_rating, $title, $publtime, $hitstotal, $author) = $result->fetch(3)) {
    $sheet->setCellValue('A' . $rowCount, $rowCount - 1);
    $sheet->setCellValue('B' . $rowCount, $title);
    $sheet->setCellValue('C' . $rowCount, $author);
    $sheet->setCellValue('D' . $rowCount, nv_date('H:i d/m/y', $publtime));
    $sheet->setCellValue('E' . $rowCount, $hitstotal);
    $sheet->setCellValue('F' . $rowCount, $total_rating);
    $rowCount++;
}

// Xuất file
$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$writer->setOffice2003Compatibility(true);
$filename=time().".xlsx";
$writer->save($filename);
header("location:".$filename);

// header('Location: ' . $_SERVER['HTTP_REFERER']);